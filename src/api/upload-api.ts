/** Outgoing attachments as the Teams web client sends them: docs/API-REFERENCE.md, Send Attachments. */

import { constants } from 'node:fs';
import { open, realpath, stat } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, extname, isAbsolute, join, sep } from 'node:path';
import { CONFIG_DIR } from '../auth/session-store.js';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError, extractRetryAfter } from '../types/errors.js';
import {
  getAmsBaseUrl,
  requireGraphTokenAsync,
  requireMessageAuthWithConfig,
} from '../utils/auth-guards.js';
import { getSkypeAuthHeaders } from '../utils/api-config.js';
import { extractOneOnOneMemberIds } from './chatsvc-messaging.js';
import { MAX_ATTACHMENT_BYTES, MAX_OUTGOING_ATTACHMENTS, MRI_ORGID_PREFIX, SELF_CHAT_ID } from '../constants.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const CHAT_FILES_FOLDER = 'Microsoft Teams Chat Files';
const UPLOAD_TIMEOUT_MS = 120_000;
const CLEANUP_DEADLINE_MS = 20_000;

const AMS_HOST = /(^|\.)asm\.skype\.com$|\.asyncgw\.teams\.microsoft\.com$/;
const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp']);

export interface OutgoingAttachments {
  imagesHtml: string;
  filesProperty?: string;
  /** Deletes uploaded files, revoking their access, and says what is left; only after a definite rejection. */
  discard: () => Promise<string>;
  sent: Array<{ name: string; kind: 'image' | 'file' }>;
}

export function attachmentKind(path: string): 'image' | 'file' {
  return IMAGE_EXTENSIONS.has(extname(path).toLowerCase()) ? 'image' : 'file';
}

export function isChannelConversation(conversationId: string): boolean {
  return /@thread\.(tacv2|skype)$/.test(conversationId);
}

/** Pixel size from PNG, GIF, WebP or JPEG headers; null for anything else. */
export function readImageSize(bytes: Uint8Array): { width: number; height: number } | null {
  const b = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (b.length >= 24 && b.readUInt32BE(0) === 0x89504e47) {
    return { width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
  }
  if (b.length >= 10 && b.toString('ascii', 0, 3) === 'GIF') {
    return { width: b.readUInt16LE(6), height: b.readUInt16LE(8) };
  }
  if (b.length >= 30 && b.toString('ascii', 0, 4) === 'RIFF' && b.toString('ascii', 8, 12) === 'WEBP') {
    const chunk = b.toString('ascii', 12, 16);
    if (chunk === 'VP8X') return { width: 1 + b.readUIntLE(24, 3), height: 1 + b.readUIntLE(27, 3) };
    if (chunk === 'VP8 ') return { width: b.readUInt16LE(26) & 0x3fff, height: b.readUInt16LE(28) & 0x3fff };
    if (chunk === 'VP8L') {
      const bits = b.readUInt32LE(21);
      return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
    }
    return null;
  }
  if (b.length >= 4 && b[0] === 0xff && b[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < b.length) {
      if (b[offset] !== 0xff) return null;
      const marker = b[offset + 1];
      const length = b.readUInt16BE(offset + 2);
      // SOF0..SOF15 carry the frame size, except DHT (C4), JPG (C8) and DAC (CC).
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: b.readUInt16BE(offset + 5), width: b.readUInt16BE(offset + 7) };
      }
      offset += 2 + length;
    }
  }
  return null;
}

export function buildAmsImageHtml(amsUrl: string, objectId: string, size: { width: number; height: number } | null): string {
  const src = `${amsUrl}/v1/objects/${objectId}/views/imgo`;
  const dims = size ? ` width="${size.width}" height="${size.height}"` : '';
  return `<p><img itemtype="http://schema.skype.com/AMSImage" src="${src}"${dims} id="${objectId}" itemscope="" href="${src}" target-src="${src}" style="vertical-align:bottom"></p>`;
}

export interface UploadedDriveItem {
  id: string;
  name: string;
  webUrl: string;
  eTag?: string;
}

export function buildFileEntry(item: UploadedDriveItem, siteUrl: string): Record<string, unknown> {
  const ext = extname(item.name).slice(1).toLowerCase();
  const uniqueId = item.eTag?.match(/\{([0-9A-Fa-f-]{36})\}/)?.[1] ?? item.id;
  return {
    '@type': 'http://schema.skype.com/File',
    version: 2,
    id: uniqueId,
    baseUrl: siteUrl,
    type: ext,
    title: item.name,
    state: 'active',
    objectUrl: item.webUrl,
    providerData: '',
    itemid: item.id,
    fileName: item.name,
    fileType: ext,
    fileInfo: { itemId: null, fileUrl: item.webUrl, siteUrl, serverRelativeUrl: '', shareUrl: item.webUrl, shareId: null },
    botFileProperties: {},
    permissionScope: 'users',
    filePreview: {},
    fileChicletState: { serviceName: 'p2p', state: 'active' },
  };
}

/** Any agent reaching this server can name a path, so dotfiles and the session store never leave. */
async function isProtectedPath(resolved: string): Promise<boolean> {
  if (resolved.split(sep).some(segment => segment.startsWith('.'))) return true;
  const home = homedir();
  const browserData = ['Google/Chrome', 'Microsoft Edge', 'BraveSoftware', 'Firefox'].map(dir => join(home, 'Library', 'Application Support', dir));
  const candidates = [CONFIG_DIR, join(home, 'Library', 'Keychains'), join(home, 'Library', 'Cookies'), ...browserData, process.env.TEAMS_EXISTING_CHROME_DIR]
    .filter((dir): dir is string => Boolean(dir));
  const dirs = await Promise.all(candidates.map(dir => realpath(dir).catch(() => dir)));
  const target = resolved.toLowerCase();
  return dirs.some(dir => {
    const base = dir.toLowerCase();
    return target === base || target.startsWith(base + sep);
  });
}

interface LocalFile {
  path: string;
  name: string;
  kind: 'image' | 'file';
}

const cannotRead = (path: string, error: unknown) => err(createError(ErrorCode.INVALID_INPUT,
  `Attachment cannot be read (${(error as NodeJS.ErrnoException).code ?? 'unknown error'}): ${path}`));
const refusedLocation = (path: string) => err(createError(ErrorCode.INVALID_INPUT,
  `Attachment is in a hidden or protected location (dotfiles, session or credential stores): ${path}`));

export async function readCheckedFile(path: string): Promise<Result<Buffer<ArrayBuffer>>> {
  let resolved: string;
  try {
    resolved = await realpath(path);
  } catch (error) {
    return cannotRead(path, error);
  }
  if (await isProtectedPath(path) || await isProtectedPath(resolved)) return refusedLocation(path);
  // Nonblocking, so a FIFO cannot stall the shared server; no final symlink is followed.
  let handle;
  try {
    handle = await open(resolved, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK);
  } catch (error) {
    return cannotRead(path, error);
  }
  try {
    const info = await handle.stat();
    if (!info.isFile()) {
      return err(createError(ErrorCode.INVALID_INPUT, `Attachment is not a regular file: ${path}`));
    }
    // Resolved again after opening: an ancestor swapped in between either shows a
    // protected location here or names a different file than the descriptor holds.
    const after = await realpath(path).catch(() => null);
    if (!after || await isProtectedPath(after)) return refusedLocation(path);
    const named = await stat(after).catch(() => null);
    if (!named || named.dev !== info.dev || named.ino !== info.ino) {
      return err(createError(ErrorCode.INVALID_INPUT, `Attachment changed while it was being checked: ${path}`));
    }
    if (info.size === 0 || info.size > MAX_ATTACHMENT_BYTES) {
      return err(createError(ErrorCode.INVALID_INPUT,
        `Attachment must be between 1 byte and ${MAX_ATTACHMENT_BYTES / (1024 * 1024)} MB: ${path}`));
    }
    const bytes = Buffer.alloc(info.size);
    const { bytesRead } = await handle.read(bytes, 0, info.size, 0);
    if (bytesRead !== info.size) {
      return err(createError(ErrorCode.INVALID_INPUT, `Attachment changed while it was being read: ${path}`));
    }
    return ok(bytes);
  } finally {
    await handle.close();
  }
}

function checkFiles(conversationId: string, paths: string[]): Result<LocalFile[]> {
  if (paths.length > MAX_OUTGOING_ATTACHMENTS) {
    return err(createError(ErrorCode.INVALID_INPUT, `At most ${MAX_OUTGOING_ATTACHMENTS} attachments per message.`));
  }
  const files: LocalFile[] = [];
  for (const path of paths) {
    if (!isAbsolute(path)) {
      return err(createError(ErrorCode.INVALID_INPUT, `Attachment path must be absolute: ${path}`));
    }
    const kind = attachmentKind(path);
    if (kind === 'file' && isChannelConversation(conversationId)) {
      return err(createError(ErrorCode.INVALID_INPUT,
        `Channels accept image attachments only (png, jpg, gif, webp); ${basename(path)} is a file. Send it in a chat, or share a link to it.`));
    }
    files.push({ path, name: basename(path), kind });
  }
  return ok(files);
}

async function uploadImage(conversationId: string, file: LocalFile): Promise<Result<string>> {
  const amsUrl = getAmsBaseUrl();
  const amsOrigin = amsUrl && URL.canParse(amsUrl) ? new URL(amsUrl) : null;
  if (!amsOrigin || amsOrigin.protocol !== 'https:' || !AMS_HOST.test(amsOrigin.hostname)) {
    return err(createError(ErrorCode.API_ERROR,
      'The Teams session does not list a media service, so images cannot be uploaded. Run teams_login with refresh: true.'));
  }
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const bytes = await readCheckedFile(file.path);
  if (!bytes.ok) return bytes;
  const ams = amsOrigin.origin;
  // AMS rejects requests without a client version it can map to a platform.
  const headers = {
    Authorization: `skype_token ${authResult.value.auth.skypeToken}`,
    'X-Ms-Client-Version': '1415/1.0.0.2025010401',
  };

  const created = await httpRequest<{ id?: string }>(`${ams}/v1/objects`, {
    method: 'POST',
    redirect: 'error',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({ type: 'pish/image', permissions: { [conversationId]: ['read'] }, filename: file.name }),
  });
  if (!created.ok) return created;
  const objectId = created.value.data.id;
  if (!objectId || !/^[A-Za-z0-9_-]+$/.test(objectId)) {
    return err(createError(ErrorCode.API_ERROR, 'The media service did not return an object id for the image.'));
  }

  const put = await httpRequest<unknown>(`${ams}/v1/objects/${encodeURIComponent(objectId)}/content/imgpsh`, {
    method: 'PUT',
    redirect: 'error',
    headers: { ...headers, 'Content-Type': 'application/octet-stream' },
    body: bytes.value,
    timeoutMs: UPLOAD_TIMEOUT_MS,
  });
  if (!put.ok) return put;

  return ok(buildAmsImageHtml(ams, objectId, readImageSize(bytes.value)));
}

async function chatMemberObjectIds(conversationId: string): Promise<Result<string[]>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;
  const self = auth.userMri.replace(MRI_ORGID_PREFIX, '');

  if (conversationId === SELF_CHAT_ID) return ok([]);
  const oneOnOne = extractOneOnOneMemberIds(conversationId);
  if (oneOnOne) {
    if (!oneOnOne.includes(self)) {
      return err(createError(ErrorCode.INVALID_INPUT, 'This 1:1 conversation id does not include you, so no file is shared through it.'));
    }
    return ok(oneOnOne.filter(id => id !== self));
  }

  const response = await httpRequest<{ members?: Array<{ id?: string }> }>(
    `${baseUrl}/api/chatsvc/${region}/v1/threads/${encodeURIComponent(conversationId)}?view=msnp24Equivalent`,
    { method: 'GET', headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl) }
  );
  if (!response.ok) return response;
  // Bots (28:) never open files; anyone else outside the organisation cannot be invited by object id.
  const people = (response.value.data.members ?? []).map(member => member.id ?? '').filter(id => !id.startsWith('28:'));
  const outsiders = people.filter(id => !id.startsWith(MRI_ORGID_PREFIX));
  if (outsiders.length > 0) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `This chat has ${outsiders.length} member(s) outside your organisation who cannot be given access to an uploaded file. Send images, or share a link instead.`));
  }
  return ok(people.map(id => id.slice(MRI_ORGID_PREFIX.length)).filter(id => id !== self));
}

export async function deleteDriveItems(auth: Record<string, string>, items: UploadedDriveItem[], budgetMs = CLEANUP_DEADLINE_MS): Promise<string[]> {
  // Plain fetch: the shared client refuses every call during a rate-limit window. The session
  // lock is held, so there is a hard deadline. Each round deletes every pending file in
  // parallel; only transient failures go to the next round, after Graph's longest Retry-After.
  const deadline = Date.now() + budgetMs;
  let pending = items;
  const failed: string[] = [];
  while (pending.length > 0 && Date.now() < deadline) {
    const timeoutMs = Math.min(10_000, deadline - Date.now());
    const responses = await Promise.all(pending.map(item => fetch(`${GRAPH}/me/drive/items/${encodeURIComponent(item.id)}`, {
      method: 'DELETE',
      headers: auth,
      signal: AbortSignal.timeout(timeoutMs),
    }).catch(() => null)));
    const retry: UploadedDriveItem[] = [];
    let waitMs = 1000;
    responses.forEach((response, index) => {
      if (response && (response.ok || response.status === 404)) return;
      if (response && response.status !== 408 && response.status !== 429 && response.status < 500) {
        failed.push(pending[index].name);
        return;
      }
      retry.push(pending[index]);
      if (response) waitMs = Math.max(waitMs, extractRetryAfter(response.headers) ?? 0);
    });
    pending = retry;
    if (pending.length > 0) await new Promise(resolve => setTimeout(resolve, Math.max(0, Math.min(waitMs, deadline - Date.now()))));
  }
  return [...failed, ...pending.map(item => item.name)];
}

function cleanupNote(left: string[]): string {
  return left.length === 0
    ? 'The uploaded files were deleted, so nobody kept access.'
    : `Could not delete ${left.join(', ')} from your OneDrive "${CHAT_FILES_FOLDER}" folder within ${CLEANUP_DEADLINE_MS / 1000}s; the chat members may still have read access until you delete them there.`;
}

async function ensureChatFilesFolder(auth: Record<string, string>): Promise<Result<void>> {
  const folder = await httpRequest<unknown>(`${GRAPH}/me/drive/root:/${encodeURIComponent(CHAT_FILES_FOLDER)}`, { method: 'GET', headers: auth });
  if (folder.ok) return ok(undefined);
  if (folder.error.code !== ErrorCode.NOT_FOUND) return folder;
  const created = await httpRequest<unknown>(`${GRAPH}/me/drive/root/children`, {
    method: 'POST',
    headers: { ...auth, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: CHAT_FILES_FOLDER, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
  });
  // 409: created meanwhile by another upload.
  return created.ok || created.error.message.includes('409') ? ok(undefined) : created;
}

interface GrantedPermission {
  roles?: string[];
  grantedToV2?: { user?: { id?: string } };
  grantedToIdentitiesV2?: Array<{ user?: { id?: string } }>;
}

export function grantsCoverRecipients(status: number, permissions: GrantedPermission[] | undefined, recipientIds: string[]): boolean {
  if (status !== 200) return false;
  const readers = new Set<string>();
  for (const permission of permissions ?? []) {
    if (!permission.roles?.includes('read')) continue;
    for (const user of [permission.grantedToV2?.user, ...(permission.grantedToIdentitiesV2 ?? []).map(identity => identity.user)]) {
      if (user?.id) readers.add(user.id.toLowerCase());
    }
  }
  return recipientIds.every(id => readers.has(id.toLowerCase()));
}

interface UploadedFiles {
  filesProperty: string;
  discard: () => Promise<string>;
}

async function uploadFiles(conversationId: string, files: LocalFile[]): Promise<Result<UploadedFiles>> {
  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const auth = { Authorization: `Bearer ${tokenResult.value}` };

  const members = await chatMemberObjectIds(conversationId);
  if (!members.ok) return members;

  const drive = await httpRequest<{ webUrl?: string }>(`${GRAPH}/me/drive?$select=webUrl`, { method: 'GET', headers: auth });
  if (!drive.ok) return drive;
  const siteUrl = (drive.value.data.webUrl ?? '').replace(/\/Documents\/?$/, '/');
  const folder = await ensureChatFilesFolder(auth);
  if (!folder.ok) return folder;

  const items: UploadedDriveItem[] = [];
  const discard = async () => cleanupNote(await deleteDriveItems(auth, items));
  const fail = async (error: { message: string }) => err(createError(ErrorCode.API_ERROR,
    `${error.message} Nothing was sent. ${await discard()}`));

  for (const file of files) {
    const bytes = await readCheckedFile(file.path);
    if (!bytes.ok) {
      if (items.length > 0) await discard();
      return bytes;
    }
    const path = `${CHAT_FILES_FOLDER}/${file.name}`.split('/').map(encodeURIComponent).join('/');
    const uploaded = await httpRequest<UploadedDriveItem>(
      `${GRAPH}/me/drive/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`,
      {
        method: 'PUT',
        headers: { ...auth, 'Content-Type': 'application/octet-stream' },
        body: bytes.value,
        timeoutMs: UPLOAD_TIMEOUT_MS,
      }
    );
    if (!uploaded.ok) return fail(uploaded.error);
    items.push(uploaded.value.data);
  }

  if (members.value.length > 0) {
    for (const item of items) {
      const invited = await httpRequest<{ value?: GrantedPermission[] }>(`${GRAPH}/me/drive/items/${encodeURIComponent(item.id)}/invite`, {
        method: 'POST',
        headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          recipients: members.value.map(objectId => ({ objectId })),
          roles: ['read'],
          requireSignIn: true,
          sendInvitation: false,
        }),
      });
      if (!invited.ok) return fail({ message: `Could not share ${item.name} with the chat members: ${invited.error.message}.` });
      if (!grantsCoverRecipients(invited.value.status, invited.value.data.value, members.value)) {
        return fail({ message: `Only part of the chat could be given read access to ${item.name} (HTTP ${invited.value.status}).` });
      }
    }
  }
  return ok({ filesProperty: JSON.stringify(items.map(item => buildFileEntry(item, siteUrl))), discard });
}

export async function prepareAttachments(conversationId: string, paths: string[]): Promise<Result<OutgoingAttachments>> {
  const checked = checkFiles(conversationId, paths);
  if (!checked.ok) return checked;

  let imagesHtml = '';
  for (const file of checked.value.filter(f => f.kind === 'image')) {
    const html = await uploadImage(conversationId, file);
    if (!html.ok) return html;
    imagesHtml += html.value;
  }

  const files = checked.value.filter(f => f.kind === 'file');
  let filesProperty: string | undefined;
  let discard = async () => '';
  if (files.length > 0) {
    const uploaded = await uploadFiles(conversationId, files);
    if (!uploaded.ok) return uploaded;
    ({ filesProperty, discard } = uploaded.value);
  }

  return ok({ imagesHtml, filesProperty, discard, sent: checked.value.map(f => ({ name: f.name, kind: f.kind })) });
}

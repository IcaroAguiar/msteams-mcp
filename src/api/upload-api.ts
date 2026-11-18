/** Outgoing attachments as the Teams web client sends them: docs/API-REFERENCE.md, Send Attachments. */

import { constants } from 'node:fs';
import { open, realpath } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, extname, isAbsolute, join, sep } from 'node:path';
import { CONFIG_DIR } from '../auth/session-store.js';
import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
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

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp']);

export interface OutgoingAttachments {
  imagesHtml: string;
  filesProperty?: string;
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
  const dirs = await Promise.all([CONFIG_DIR, join(homedir(), 'Library', 'Keychains')]
    .map(dir => realpath(dir).catch(() => dir)));
  const target = resolved.toLowerCase();
  return dirs.some(dir => {
    const base = dir.toLowerCase();
    return target === base || target.startsWith(base + sep);
  });
}

interface LocalFile {
  bytes: Buffer<ArrayBuffer>;
  name: string;
  kind: 'image' | 'file';
}

async function checkFiles(conversationId: string, paths: string[]): Promise<Result<LocalFile[]>> {
  if (paths.length > MAX_OUTGOING_ATTACHMENTS) {
    return err(createError(ErrorCode.INVALID_INPUT, `At most ${MAX_OUTGOING_ATTACHMENTS} attachments per message.`));
  }
  const files: LocalFile[] = [];
  for (const path of paths) {
    if (!isAbsolute(path)) {
      return err(createError(ErrorCode.INVALID_INPUT, `Attachment path must be absolute: ${path}`));
    }
    const cannotRead = (error: unknown) => err(createError(ErrorCode.INVALID_INPUT,
      `Attachment cannot be read (${(error as NodeJS.ErrnoException).code ?? 'unknown error'}): ${path}`));
    let resolved: string;
    try {
      resolved = await realpath(path);
    } catch (error) {
      return cannotRead(error);
    }
    if (await isProtectedPath(resolved)) {
      return err(createError(ErrorCode.INVALID_INPUT,
        `Attachment is in a hidden or protected location (dotfiles, session or credential stores): ${path}`));
    }
    // The checked descriptor is the one read, so a swapped symlink cannot redirect the upload.
    let handle;
    try {
      handle = await open(resolved, constants.O_RDONLY | constants.O_NOFOLLOW);
    } catch (error) {
      return cannotRead(error);
    }
    let bytes: Buffer<ArrayBuffer>;
    try {
      const info = await handle.stat();
      if (!info.isFile()) {
        return err(createError(ErrorCode.INVALID_INPUT, `Attachment is not a regular file: ${path}`));
      }
      if (info.size === 0 || info.size > MAX_ATTACHMENT_BYTES) {
        return err(createError(ErrorCode.INVALID_INPUT,
          `Attachment must be between 1 byte and ${MAX_ATTACHMENT_BYTES / (1024 * 1024)} MB: ${path}`));
      }
      bytes = await handle.readFile();
    } finally {
      await handle.close();
    }
    const kind = attachmentKind(path);
    if (kind === 'file' && isChannelConversation(conversationId)) {
      return err(createError(ErrorCode.INVALID_INPUT,
        `Channels accept image attachments only (png, jpg, gif, webp); ${basename(path)} is a file. Send it in a chat, or share a link to it.`));
    }
    files.push({ bytes, name: basename(path), kind });
  }
  return ok(files);
}

async function uploadImage(conversationId: string, file: LocalFile): Promise<Result<string>> {
  const amsUrl = getAmsBaseUrl();
  if (!amsUrl) {
    return err(createError(ErrorCode.API_ERROR,
      'The Teams session does not list a media service, so images cannot be uploaded. Run teams_login with refresh: true.'));
  }
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, baseUrl } = authResult.value;
  // AMS rejects requests without a client version it can map to a platform.
  const headers = {
    ...getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    Authorization: `skype_token ${auth.skypeToken}`,
    'X-Ms-Client-Version': '1415/1.0.0.2025010401',
  };

  const created = await httpRequest<{ id?: string }>(`${amsUrl}/v1/objects`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'application/json' },
    body: JSON.stringify({ type: 'pish/image', permissions: { [conversationId]: ['read'] }, filename: file.name }),
  });
  if (!created.ok) return created;
  const objectId = created.value.data.id;
  if (!objectId) {
    return err(createError(ErrorCode.API_ERROR, 'The media service did not return an object id for the image.'));
  }

  const bytes = file.bytes;
  const put = await httpRequest<unknown>(`${amsUrl}/v1/objects/${encodeURIComponent(objectId)}/content/imgpsh`, {
    method: 'PUT',
    headers: { ...headers, 'Content-Type': 'application/octet-stream' },
    body: bytes,
    timeoutMs: UPLOAD_TIMEOUT_MS,
  });
  if (!put.ok) return put;

  return ok(buildAmsImageHtml(amsUrl, objectId, readImageSize(bytes)));
}

async function chatMemberObjectIds(conversationId: string): Promise<Result<string[]>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;
  const self = auth.userMri.replace(MRI_ORGID_PREFIX, '');

  if (conversationId === SELF_CHAT_ID) return ok([]);
  const oneOnOne = extractOneOnOneMemberIds(conversationId);
  if (oneOnOne) return ok(oneOnOne.filter(id => id !== self));

  const response = await httpRequest<{ members?: Array<{ id?: string }> }>(
    `${baseUrl}/api/chatsvc/${region}/v1/threads/${encodeURIComponent(conversationId)}?view=msnp24Equivalent`,
    { method: 'GET', headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl) }
  );
  if (!response.ok) return response;
  const ids = (response.value.data.members ?? [])
    .map(member => member.id ?? '')
    .filter(id => id.startsWith(MRI_ORGID_PREFIX))
    .map(id => id.slice(MRI_ORGID_PREFIX.length))
    .filter(id => id !== self);
  return ok(ids);
}

async function uploadFiles(conversationId: string, files: LocalFile[]): Promise<Result<string>> {
  const tokenResult = await requireGraphTokenAsync();
  if (!tokenResult.ok) return tokenResult;
  const auth = { Authorization: `Bearer ${tokenResult.value}` };

  const members = await chatMemberObjectIds(conversationId);
  if (!members.ok) return members;

  const drive = await httpRequest<{ webUrl?: string }>(`${GRAPH}/me/drive?$select=webUrl`, { method: 'GET', headers: auth });
  if (!drive.ok) return drive;
  const siteUrl = (drive.value.data.webUrl ?? '').replace(/\/Documents\/?$/, '/');

  const entries: Record<string, unknown>[] = [];
  for (const file of files) {
    const path = `${CHAT_FILES_FOLDER}/${file.name}`.split('/').map(encodeURIComponent).join('/');
    const uploaded = await httpRequest<UploadedDriveItem>(
      `${GRAPH}/me/drive/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`,
      {
        method: 'PUT',
        headers: { ...auth, 'Content-Type': 'application/octet-stream' },
        body: file.bytes,
        timeoutMs: UPLOAD_TIMEOUT_MS,
      }
    );
    if (!uploaded.ok) return uploaded;
    const item = uploaded.value.data;

    if (members.value.length > 0) {
      const invited = await httpRequest<unknown>(`${GRAPH}/me/drive/items/${encodeURIComponent(item.id)}/invite`, {
        method: 'POST',
        headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          recipients: members.value.map(objectId => ({ objectId })),
          roles: ['read'],
          requireSignIn: true,
          sendInvitation: false,
        }),
      });
      if (!invited.ok) {
        return err(createError(ErrorCode.API_ERROR,
          `Uploaded ${item.name} to your OneDrive but could not share it with the chat members: ${invited.error.message}. Nothing was sent.`));
      }
    }
    entries.push(buildFileEntry(item, siteUrl));
  }
  return ok(JSON.stringify(entries));
}

/** Uploads everything before the message is sent, so a failure sends nothing. */
export async function prepareAttachments(conversationId: string, paths: string[]): Promise<Result<OutgoingAttachments>> {
  const checked = await checkFiles(conversationId, paths);
  if (!checked.ok) return checked;

  let imagesHtml = '';
  for (const file of checked.value.filter(f => f.kind === 'image')) {
    const html = await uploadImage(conversationId, file);
    if (!html.ok) return html;
    imagesHtml += html.value;
  }

  const files = checked.value.filter(f => f.kind === 'file');
  let filesProperty: string | undefined;
  if (files.length > 0) {
    const uploaded = await uploadFiles(conversationId, files);
    if (!uploaded.ok) return uploaded;
    filesProperty = uploaded.value;
  }

  return ok({ imagesHtml, filesProperty, sent: checked.value.map(f => ({ name: f.name, kind: f.kind })) });
}

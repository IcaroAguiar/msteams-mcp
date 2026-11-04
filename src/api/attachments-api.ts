import { createHash, randomBytes } from 'node:crypto';
import { chmod, mkdir, readdir, rename, unlink } from 'node:fs/promises';
import * as path from 'node:path';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { CONFIG_DIR, ensureConfigDir } from '../auth/session-store.js';
import { MAX_ATTACHMENT_BYTES } from '../constants.js';
import { requireMessageAuth } from '../utils/auth-guards.js';
import { isTeamsMediaUrl, type MessageAttachment } from '../utils/parsers-message-parts.js';
import { getMessage } from './chatsvc-messaging.js';
import { downloadFile, streamToFile } from './graph-files-api.js';

export const ATTACHMENTS_DIR = path.join(CONFIG_DIR, 'attachments');

const DOWNLOAD_ATTEMPTS = 3;
const MAX_NAME_LENGTH = 120;
const RETRY_BASE_DELAY_MS = 500;

export interface SavedAttachment {
  id: string;
  kind: MessageAttachment['kind'];
  name?: string;
  path?: string;
  size?: number;
  contentType?: string;
  sha256?: string;
  cached?: boolean;
  error?: { code: ErrorCode; message: string };
}

export interface SaveAttachmentsResult {
  conversationId: string;
  messageId: string;
  directory: string;
  attachments: SavedAttachment[];
}

const EXTENSIONS: Record<string, string> = {
  'image/jpeg': 'jpg', 'image/png': 'png', 'image/gif': 'gif', 'image/webp': 'webp',
  'image/bmp': 'bmp', 'image/heic': 'heic', 'image/svg+xml': 'svg', 'video/mp4': 'mp4',
};

function shortHash(value: string): string {
  return createHash('sha256').update(value).digest('hex').slice(0, 12);
}

export function safeFileName(name: string): string {
  const cleaned = path.basename(name.normalize('NFC'))
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f\u007f/\\:*?"<>|]/g, '_')
    .replace(/^\.+/, '')
    .trim() || 'file';
  if (cleaned.length <= MAX_NAME_LENGTH) return cleaned;
  const extension = path.extname(cleaned).slice(0, 16);
  return cleaned.slice(0, MAX_NAME_LENGTH - extension.length) + extension;
}

// Conversation IDs are hashed, never used as path segments.
async function messageDirectory(conversationId: string, messageId: string): Promise<string> {
  ensureConfigDir();
  const conversationDir = path.join(ATTACHMENTS_DIR, shortHash(conversationId));
  const directory = path.join(conversationDir, messageId);
  for (const dir of [ATTACHMENTS_DIR, conversationDir, directory]) {
    await mkdir(dir, { recursive: true, mode: 0o700 });
    await chmod(dir, 0o700);
  }
  return directory;
}

function fileStem(attachment: MessageAttachment): string {
  const label = attachment.kind === 'image' && !attachment.name ? 'image' : safeFileName(attachment.name ?? attachment.kind);
  return `${shortHash(attachment.id)}-${label}`;
}

async function existingFile(directory: string, stem: string): Promise<string | undefined> {
  const names = await readdir(directory);
  return names.find(name => name === stem || name.startsWith(`${stem}.`));
}

async function withRetry<T>(attempt: () => Promise<Result<T>>): Promise<Result<T>> {
  let result = await attempt();
  for (let n = 1; n < DOWNLOAD_ATTEMPTS && !result.ok && result.error.retryable; n++) {
    await new Promise(resolve => setTimeout(resolve, RETRY_BASE_DELAY_MS * 2 ** (n - 1)));
    result = await attempt();
  }
  return result;
}

async function saveOne(attachment: MessageAttachment, directory: string, skypeToken: string): Promise<Result<SavedAttachment>> {
  const stem = fileStem(attachment);
  const found = await existingFile(directory, stem);
  if (found) {
    return ok({ id: attachment.id, kind: attachment.kind, name: attachment.name, path: path.join(directory, found), cached: true });
  }

  const partial = path.join(directory, `.${stem}.${randomBytes(6).toString('hex')}.partial`);
  let saved: Result<{ size: number; sha256: string; contentType?: string }>;
  if (attachment.source === 'teams-media') {
    if (!isTeamsMediaUrl(attachment.url)) {
      return err(createError(ErrorCode.INVALID_INPUT, 'The image is not stored in Teams media storage', { retryable: false }));
    }
    saved = await withRetry(() => streamToFile(attachment.url, partial, undefined, { Authorization: `skype_token ${skypeToken}` }, MAX_ATTACHMENT_BYTES));
  } else {
    saved = await withRetry(() => downloadFile(attachment.url, partial, MAX_ATTACHMENT_BYTES));
  }
  if (!saved.ok) return saved;

  const served = saved.value.contentType?.split(';')[0].trim().toLowerCase();
  const reported = /^[a-z0-9]{1,10}$/i.test(attachment.type ?? '') ? attachment.type : undefined;
  const extension = path.extname(stem) ? '' : `.${EXTENSIONS[served ?? ''] ?? reported ?? 'bin'}`;
  const finalPath = path.join(directory, `${stem}${extension}`);
  try {
    await rename(partial, finalPath);
  } catch (error) {
    await unlink(partial).catch(() => {});
    return err(createError(ErrorCode.UNKNOWN, `Could not save the attachment: ${error instanceof Error ? error.message : String(error)}`));
  }
  return ok({
    id: attachment.id,
    kind: attachment.kind,
    name: attachment.name,
    path: finalPath,
    size: saved.value.size,
    contentType: served,
    sha256: saved.value.sha256,
  });
}

export async function saveMessageAttachments(
  conversationId: string,
  messageId: string,
  attachmentId?: string,
): Promise<Result<SaveAttachmentsResult>> {
  if (!/^\d+$/.test(messageId)) {
    return err(createError(ErrorCode.INVALID_INPUT, 'messageId must be the numeric message ID from teams_get_thread or teams_get_message'));
  }
  const message = await getMessage(conversationId, messageId);
  if (!message.ok) return message;

  const all = message.value.attachments ?? [];
  const wanted = attachmentId ? all.filter(a => a.id === attachmentId) : all;
  if (wanted.length === 0) {
    return err(createError(ErrorCode.NOT_FOUND,
      attachmentId ? `Attachment ${attachmentId} is not in this message` : 'This message has no images or files',
      { retryable: false, suggestions: ['Use the attachments listed by teams_get_message for this message'] }));
  }

  const auth = requireMessageAuth();
  if (!auth.ok) return auth;
  const directory = await messageDirectory(conversationId, messageId);

  const attachments: SavedAttachment[] = [];
  const errors: McpError[] = [];
  for (const attachment of wanted) {
    const saved = await saveOne(attachment, directory, auth.value.skypeToken);
    if (saved.ok) {
      attachments.push(saved.value);
    } else {
      errors.push(saved.error);
      attachments.push({ id: attachment.id, kind: attachment.kind, name: attachment.name,
        error: { code: saved.error.code, message: saved.error.message } });
    }
  }
  // An expired token must reach the server's renewal; files already saved are reused on the retry.
  const expired = errors.find(e => e.code === ErrorCode.AUTH_EXPIRED);
  if (expired) return err(expired);
  if (errors.length > 0 && attachments.every(a => a.error)) return err(errors[0]);
  return ok({ conversationId, messageId, directory, attachments });
}

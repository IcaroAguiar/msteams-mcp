import { afterAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { chmodSync, mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { readFile, readdir, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import type { ThreadMessage } from './chatsvc-messaging.js';
import type { MessageAttachment } from '../utils/parsers-message-parts.js';

const root = vi.hoisted(() => ({ dir: '' }));
root.dir = mkdtempSync(join(tmpdir(), 'teams-attachments-'));

vi.mock('../auth/session-store.js', () => ({
  get CONFIG_DIR() { return root.dir; },
  ensureConfigDir: () => mkdirSync(root.dir, { recursive: true, mode: 0o700 }),
}));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuth: () => ({ ok: true, value: { skypeToken: 'skype-token' } }),
}));
vi.mock('./chatsvc-messaging.js', () => ({ getMessage: vi.fn() }));
vi.mock('./graph-files-api.js', async (importOriginal) => ({
  ...await importOriginal<typeof import('./graph-files-api.js')>(),
  downloadFile: vi.fn(),
}));

const { saveMessageAttachments, safeFileName, ATTACHMENTS_DIR } = await import('./attachments-api.js');
const { getMessage } = await import('./chatsvc-messaging.js');
const { downloadFile } = await import('./graph-files-api.js');

const CONVERSATION = '19:synthetic-chat@unq.gbl.spaces';
const IMAGE: MessageAttachment = {
  id: '0-wus-d1-0123456789abcdef', kind: 'image', source: 'teams-media', type: 'png',
  url: 'https://us-api.asm.skype.com/v1/objects/0-wus-d1-0123456789abcdef/views/imgo',
};
const FILE: MessageAttachment = {
  id: 'file-a', kind: 'file', source: 'sharepoint', name: 'notes.md', type: 'md',
  url: 'https://contoso-my.sharepoint.com/personal/someone/Documents/notes.md',
};
const JPEG = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]);

function message(attachments: MessageAttachment[]): ThreadMessage {
  return {
    id: '1700000000000', content: '', contentType: 'RichText/Html', sender: { mri: '8:orgid:x' },
    timestamp: '2026-01-01T00:00:00Z', conversationId: CONVERSATION, attachments,
  };
}

const fetchMock = vi.fn();

beforeEach(() => {
  vi.clearAllMocks();
  rmSync(root.dir, { recursive: true, force: true });
  fetchMock.mockImplementation(async () => new Response(JPEG, { status: 200, headers: { 'content-type': 'image/jpeg' } }));
  vi.stubGlobal('fetch', fetchMock);
  vi.mocked(downloadFile).mockImplementation(async (_url, outputPath) => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(outputPath, '# notes', { flag: 'wx', mode: 0o600 });
    return ok({ name: 'notes.md', outputPath, size: 7, contentType: 'text/markdown', sha256: 'abc' });
  });
});

afterAll(() => rmSync(root.dir, { recursive: true, force: true }));

describe('saveMessageAttachments', () => {
  it('saves a pasted image privately, named by its served type, sending the Skype token only to media storage', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));

    const result = await saveMessageAttachments(CONVERSATION, '1700000000000');

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const [saved] = result.value.attachments;
    expect(saved).toMatchObject({ id: IMAGE.id, kind: 'image', contentType: 'image/jpeg', size: JPEG.length });
    expect(saved.path).toMatch(/-image\.jpg$/);
    expect(await readFile(saved.path!)).toEqual(JPEG);
    expect((await stat(saved.path!)).mode & 0o777).toBe(0o600);
    for (const dir of [ATTACHMENTS_DIR, join(result.value.directory, '..'), result.value.directory]) {
      expect((await stat(dir)).mode & 0o777).toBe(0o700);
    }
    expect(result.value.directory).not.toContain(CONVERSATION);
    expect(fetchMock).toHaveBeenCalledWith(IMAGE.url, expect.objectContaining({
      headers: { Authorization: 'skype_token skype-token' }, redirect: 'error',
    }));
    expect((await readdir(result.value.directory)).filter(name => name.endsWith('.partial'))).toEqual([]);
  });

  it('reuses an earlier download instead of fetching again', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));
    const first = await saveMessageAttachments(CONVERSATION, '1700000000000');
    const second = await saveMessageAttachments(CONVERSATION, '1700000000000');

    expect(second).toMatchObject({ ok: true, value: { attachments: [{ cached: true }] } });
    if (first.ok && second.ok) expect(second.value.attachments[0].path).toBe(first.value.attachments[0].path);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('retries a transient media failure', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));
    fetchMock.mockRejectedValueOnce(new TypeError('fetch failed'));

    const result = await saveMessageAttachments(CONVERSATION, '1700000000000');

    expect(result).toMatchObject({ ok: true, value: { attachments: [{ contentType: 'image/jpeg' }] } });
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('reports an expired Skype token so the server can renew it', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));
    fetchMock.mockResolvedValue(new Response('', { status: 401 }));

    expect(await saveMessageAttachments(CONVERSATION, '1700000000000'))
      .toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });
  });

  it('reports an expired token even when another attachment was saved or failed first', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([FILE, IMAGE])));
    fetchMock.mockResolvedValue(new Response('', { status: 401 }));
    expect(await saveMessageAttachments(CONVERSATION, '1700000000000'))
      .toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });

    const second = { ...IMAGE, id: '0-wus-d1-fedcba9876543210', url: 'https://us-api.asm.skype.com/v1/objects/0-wus-d1-fedcba9876543210/views/imgo' };
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE, second])));
    fetchMock.mockResolvedValueOnce(new Response('', { status: 404 })).mockResolvedValue(new Response('', { status: 401 }));
    expect(await saveMessageAttachments(CONVERSATION, '1700000000001'))
      .toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });
  });

  it('downloads SharePoint files through Graph and saves the rest when one attachment fails', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE, FILE])));
    fetchMock.mockResolvedValue(new Response('', { status: 404 }));

    const result = await saveMessageAttachments(CONVERSATION, '1700000000000');

    expect(downloadFile).toHaveBeenCalledWith(FILE.url, expect.stringMatching(/\.partial$/), 100 * 1024 * 1024);
    expect(result).toMatchObject({ ok: true, value: { attachments: [
      { id: IMAGE.id, error: { code: ErrorCode.NOT_FOUND } },
      { id: FILE.id, name: 'notes.md', contentType: 'text/markdown' },
    ] } });
    if (result.ok) expect(result.value.attachments[1].path).toMatch(/-notes\.md$/);
  });

  it('saves only the requested attachment', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE, FILE])));

    const result = await saveMessageAttachments(CONVERSATION, '1700000000000', FILE.id);

    expect(result).toMatchObject({ ok: true, value: { attachments: [{ id: FILE.id }] } });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('refuses media URLs outside Teams media storage', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([{ ...IMAGE, url: 'https://attacker.example/v1/objects/x' }])));

    expect(await saveMessageAttachments(CONVERSATION, '1700000000000'))
      .toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('explains a message without attachments and rejects non-numeric IDs', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([])));
    expect(await saveMessageAttachments(CONVERSATION, '1700000000000')).toMatchObject({ ok: false, error: { code: ErrorCode.NOT_FOUND } });
    expect(await saveMessageAttachments(CONVERSATION, '../1')).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
  });

  it('passes message lookup errors through', async () => {
    vi.mocked(getMessage).mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'expired')));
    expect(await saveMessageAttachments(CONVERSATION, '1700000000000')).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });
  });
});

describe('safeFileName', () => {
  it('keeps names inside their directory and visible', () => {
    expect(safeFileName('../../etc/passwd')).toBe('passwd');
    expect(safeFileName('.hidden')).toBe('hidden');
    expect(safeFileName('a\u0000b:c?.txt')).toBe('a_b_c_.txt');
    expect(safeFileName('')).toBe('file');
  });

  it('shortens long names but keeps the extension', () => {
    const name = safeFileName(`${'x'.repeat(300)}.pdf`);
    expect(name).toHaveLength(120);
    expect(name.endsWith('.pdf')).toBe(true);
  });
});

describe('private storage', () => {
  const messageId = '1700000000000';

  it('refuses to chmod or write through a symlinked attachments folder', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));
    const outside = mkdtempSync(join(tmpdir(), 'teams-outside-'));
    chmodSync(outside, 0o755);
    mkdirSync(root.dir, { recursive: true });
    symlinkSync(outside, ATTACHMENTS_DIR);

    const result = await saveMessageAttachments(CONVERSATION, messageId);

    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.UNKNOWN } });
    expect((await stat(outside)).mode & 0o777).toBe(0o755);
    expect(await readdir(outside)).toEqual([]);
    rmSync(outside, { recursive: true, force: true });
  });

  it('does not reuse a cached entry that is a symlink, group-readable or over the size limit', async () => {
    vi.mocked(getMessage).mockResolvedValue(ok(message([IMAGE])));
    const first = await saveMessageAttachments(CONVERSATION, messageId);
    if (!first.ok) throw new Error('setup failed');
    const saved = first.value.attachments[0].path!;

    chmodSync(saved, 0o644);
    const loose = await saveMessageAttachments(CONVERSATION, messageId);
    expect(loose.ok && loose.value.attachments[0].cached).toBeFalsy();
    expect((await stat(saved)).mode & 0o777).toBe(0o600);

    rmSync(saved);
    const target = join(root.dir, 'elsewhere');
    writeFileSync(target, 'x');
    symlinkSync(target, saved);
    const linked = await saveMessageAttachments(CONVERSATION, messageId);
    expect(linked.ok && linked.value.attachments[0].cached).toBeFalsy();
    expect(await readFile(target, 'utf8')).toBe('x');
  });
});

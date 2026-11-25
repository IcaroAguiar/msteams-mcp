import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { httpRequest, clearRateLimitState } from '../utils/http.js';
import { requireGraphTokenAsync, requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { getConversationProperties } from './chatsvc-messaging.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { classifyTab, listConversationTabs, readLoopPage } from './tabs-api.js';
import { MAX_NOTES_TEXT_CHARS } from '../constants.js';

vi.mock('../utils/http.js', async (original) => ({
  ...await original<typeof import('../utils/http.js')>(),
  httpRequest: vi.fn(),
}));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphTokenAsync: vi.fn(), requireMessageAuthWithConfig: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ invalidateAccessToken: vi.fn() }));
vi.mock('./chatsvc-messaging.js', async (original) => ({
  ...await original<typeof import('./chatsvc-messaging.js')>(),
  getConversationProperties: vi.fn(),
}));

const mockHttp = vi.mocked(httpRequest);
const CHAT = '19:abc@thread.v2';
const CHANNEL = '19:def@thread.tacv2';
const NOTES_URL = 'https://contoso.sharepoint.com/contentstorage/CSP_1/Document%20Library/19_abc_thread_v2/loop_loopThread.loop?nav=x';
const PAGE_URL = 'https://contoso.sharepoint.com/:fl:/r/sites/team/Shared%20Documents/General/Plan.loop?d=w1&csf=1';

const response = <T>(data: T, headers: Record<string, string> = {}, status = 200) => ok({ status, headers: new Headers(headers), data });
const html = (data: string) => response(data, { 'content-type': 'text/html; charset=utf-8' });
const fileAdd = (subtype: string) => ({
  messagetype: 'ThreadActivity/FileAdd',
  content: `<fileadd><target><id>t-${subtype}</id><uri>${NOTES_URL.replace(/&/g, '&amp;')}</uri><type>loop</type><subtype>${subtype}</subtype></target></fileadd>`,
});

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(requireGraphTokenAsync).mockResolvedValue(ok('graph-token'));
  vi.mocked(requireMessageAuthWithConfig).mockReturnValue(ok({
    auth: { skypeToken: 'skype', authToken: 'auth', userMri: '8:orgid:me' },
    region: 'amer',
    baseUrl: 'https://teams.microsoft.com',
  } as never));
});

describe('classifyTab', () => {
  it('marks Loop page tabs readable and explains where other tab content lives', () => {
    expect(classifyTab({ id: '1', displayName: 'Page', configuration: { contentUrl: PAGE_URL }, teamsApp: { id: '5ae80e49-7ada-461a-a6bd-c5df2e0cdb06' } }))
      .toMatchObject({ kind: 'loop', readable: true, contentUrl: PAGE_URL });
    expect(classifyTab({ id: '2', displayName: 'Notes', configuration: { contentUrl: 'https://www.onenote.com/teams/ChannelNotes' }, teamsApp: { id: '0d820ecd-def2-4297-adad-78056cde7c78' } }))
      .toMatchObject({ kind: 'onenote', readable: false, reason: expect.stringContaining('OneNote') });
    expect(classifyTab({ id: '3', displayName: 'Site', configuration: { contentUrl: 'https://example.com/x.loop' }, teamsApp: { id: 'com.microsoft.teamspace.tab.web' } }))
      .toMatchObject({ kind: 'website', readable: false });
  });
});

describe('listConversationTabs', () => {
  it('finds a chat\'s Notes page in its history and lists the Graph tabs', async () => {
    vi.mocked(getConversationProperties).mockResolvedValue(ok({ conversationType: 'Chat' }));
    mockHttp.mockImplementation(async (url) => {
      if (url.startsWith('https://graph.microsoft.com')) {
        return response({ value: [{ id: 'g1', displayName: 'Plan', configuration: { contentUrl: PAGE_URL }, teamsApp: { id: 'x' } }] });
      }
      return response({ messages: [{ messagetype: 'RichText/Html', content: 'hi' }, fileAdd('page'), fileAdd('loopthread')] });
    });

    const result = await listConversationTabs(CHAT);

    expect(result.ok && result.value).toMatchObject({
      conversationType: 'chat',
      notesSearch: { found: true, messagesScanned: 3 },
      tabs: [
        { id: 'notes:t-loopthread', kind: 'notes', readable: true, contentUrl: NOTES_URL },
        { id: 'g1', kind: 'loop', readable: true },
      ],
    });
    expect(String(mockHttp.mock.calls.find(([u]) => u.startsWith('https://graph'))?.[0]))
      .toBe(`https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CHAT)}/tabs?$expand=teamsApp`);
  });

  it('pages back through history and reports a search that reached the start without Notes', async () => {
    vi.mocked(getConversationProperties).mockResolvedValue(ok({ conversationType: 'Chat' }));
    let chatsvcCalls = 0;
    mockHttp.mockImplementation(async (url) => {
      if (url.startsWith('https://graph.microsoft.com')) return response({ value: [] });
      chatsvcCalls++;
      return chatsvcCalls === 1
        ? response({ messages: [{ messagetype: 'Text' }], _metadata: { backwardLink: 'https://x/messages?startTime=1&syncState=abc' } })
        : response({ messages: [{ messagetype: 'Text' }] });
    });

    const result = await listConversationTabs(CHAT);

    expect(result.ok && result.value).toMatchObject({ tabs: [], notesSearch: { found: false, reachedStart: true, messagesScanned: 2 } });
    expect(chatsvcCalls).toBe(2);
  });

  it('keeps the Notes tab when Graph cannot list app tabs, without asking to sign in again', async () => {
    vi.mocked(getConversationProperties).mockResolvedValue(ok({ conversationType: 'Chat' }));
    mockHttp.mockImplementation(async (url) => url.startsWith('https://graph.microsoft.com')
      ? err(createError(ErrorCode.AUTH_REQUIRED, 'HTTP 403'))
      : response({ messages: [fileAdd('loopthread')] }));

    const result = await listConversationTabs(CHAT);

    expect(result.ok && result.value).toMatchObject({
      tabs: [{ kind: 'notes' }],
      tabsError: { code: ErrorCode.ACCESS_DENIED, retryable: false },
    });
  });

  describe('when Graph answers 429', () => {
    afterEach(() => {
      clearRateLimitState();
      vi.unstubAllGlobals();
    });

    it('still returns the Notes tab with tabsError, and the throttle holds back Graph only', async () => {
      const { httpRequest: realHttpRequest } = await vi.importActual<typeof import('../utils/http.js')>('../utils/http.js');
      mockHttp.mockImplementation(realHttpRequest);
      vi.mocked(getConversationProperties).mockResolvedValue(ok({ conversationType: 'Chat' }));
      const fetchMock = vi.fn(async (url: string) => url.startsWith('https://graph.microsoft.com')
        ? new Response('Too Many Requests', { status: 429, headers: { 'Retry-After': '60' } })
        : new Response(JSON.stringify({ messages: [fileAdd('loopthread')] }), { headers: { 'Content-Type': 'application/json' } }));
      vi.stubGlobal('fetch', fetchMock);

      for (let call = 0; call < 2; call++) {
        const result = await listConversationTabs(CHAT);
        expect(result.ok && result.value).toMatchObject({
          tabs: [{ kind: 'notes' }],
          tabsError: { code: ErrorCode.RATE_LIMITED, retryable: true },
        });
      }
      const graphCalls = fetchMock.mock.calls.filter(([url]) => url.startsWith('https://graph.microsoft.com'));
      expect(graphCalls).toHaveLength(1);
      expect(fetchMock).toHaveBeenCalledTimes(3);
    });
  });

  it('reads channel tabs through the team and does not search channel history', async () => {
    vi.mocked(getConversationProperties).mockResolvedValue(ok({ conversationType: 'Channel', groupId: 'group-1' }));
    mockHttp.mockResolvedValue(response({ value: [] }));

    const result = await listConversationTabs(`${CHANNEL};messageid=123`);

    expect(result.ok && result.value).toMatchObject({ conversationId: CHANNEL, conversationType: 'channel', tabs: [] });
    expect(mockHttp).toHaveBeenCalledTimes(1);
    expect(mockHttp.mock.calls[0][0]).toBe(`https://graph.microsoft.com/v1.0/teams/group-1/channels/${encodeURIComponent(CHANNEL)}/tabs?$expand=teamsApp`);
  });
});

describe('readLoopPage', () => {
  const item = { id: 'item-1', name: 'loop_loopThread.loop', file: {}, parentReference: { driveId: 'b!drive' }, lastModifiedBy: { user: { displayName: 'Alex' } } };

  it('renders the Loop file through Graph and returns its text', async () => {
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(html('<h2>Title</h2><ul><li>One</li><li>Two</li></ul>'));

    const result = await readLoopPage(NOTES_URL);

    expect(result).toEqual(ok(expect.objectContaining({ text: '## Title\n- One\n- Two', lineCount: 3, truncated: false, lastModifiedBy: 'Alex' })));
    const [metaUrl] = mockHttp.mock.calls[0];
    expect(new URL(metaUrl).pathname).toBe(`/v1.0/shares/u!${Buffer.from(NOTES_URL).toString('base64url')}/driveItem`);
    expect(mockHttp.mock.calls[1][0]).toBe('https://graph.microsoft.com/v1.0/drives/b!drive/items/item-1/content?format=html');
    for (const [, options] of mockHttp.mock.calls) {
      expect((options as { method: string }).method).toBe('GET');
    }
  });

  it('refuses URLs outside SharePoint and Graph before any request', async () => {
    const result = await readLoopPage('https://evil.example.com/notes.loop');
    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('refuses files that are not Loop pages', async () => {
    mockHttp.mockResolvedValueOnce(response({ ...item, name: 'Budget.xlsx' }));
    const result = await readLoopPage('https://contoso.sharepoint.com/sites/team/Budget.xlsx');
    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT, message: expect.stringContaining('teams_download_file') } });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('accepts only an HTML rendering, and keeps the Graph token when the storage link is refused', async () => {
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(response('plain text', { 'content-type': 'text/plain' }));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, message: expect.stringContaining('text/plain') } });

    const storage = 'https://contoso.sharepoint.com/_layouts/15/download.aspx?tempauth=x';
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(response('', { location: storage }, 302))
      .mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: true } });
    const [storageUrl, storageOptions] = mockHttp.mock.calls[4];
    expect(storageUrl).toBe(storage);
    expect((storageOptions as { headers: Record<string, string> }).headers).not.toHaveProperty('Authorization');
    expect(invalidateAccessToken).not.toHaveBeenCalled();
  });

  it('follows the storage redirect only to SharePoint or OneDrive hosts, one hop at a time', async () => {
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(response('', { location: 'https://evil.example.com/page.html' }, 302));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, message: expect.stringContaining('unexpected location') } });
    expect(mockHttp).toHaveBeenCalledTimes(2);

    mockHttp.mockClear();
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(response('', { location: 'https://region1-mediap.svc.ms/transform/x' }, 302))
      .mockResolvedValueOnce(response('', { location: 'https://evil.sharepoint.com.example.net/x' }, 302));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: false, error: { message: expect.stringContaining('unexpected location') } });
    expect(mockHttp).toHaveBeenCalledTimes(3);
    expect(mockHttp.mock.calls[2][1]).toMatchObject({ redirect: 'manual' });
  });

  it('accepts text/html with parameters but not a type that only starts with text/html', async () => {
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(html('<p>Hi</p>'));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: true, value: { text: 'Hi' } });

    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(response('<p>Hi</p>', { 'content-type': 'text/html+json' }));
    expect(await readLoopPage(NOTES_URL)).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
  });

  it('does not split an emoji at the text limit', async () => {
    mockHttp
      .mockResolvedValueOnce(response(item))
      .mockResolvedValueOnce(html(`<p>${'a'.repeat(MAX_NOTES_TEXT_CHARS - 1)}\u{1F600}b</p>`));
    const result = await readLoopPage(NOTES_URL);
    expect(result).toMatchObject({ ok: true, value: { truncated: true, text: 'a'.repeat(MAX_NOTES_TEXT_CHARS - 1) } });
  });

  it('discards a token Graph rejects and never reports AUTH_EXPIRED', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    const result = await readLoopPage(NOTES_URL);
    expect(result).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: true } });
    expect(invalidateAccessToken).toHaveBeenCalledWith('graph-token');
  });
});

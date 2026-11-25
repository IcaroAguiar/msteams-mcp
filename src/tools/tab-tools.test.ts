import { describe, expect, it, vi } from 'vitest';
import { GetNotesInputSchema, getNotesTool } from './index.js';
import { listConversationTabs } from '../api/tabs-api.js';
import { ok } from '../types/result.js';
import { ErrorCode } from '../types/errors.js';

vi.mock('../api/tabs-api.js', () => ({ listConversationTabs: vi.fn(), readLoopPage: vi.fn() }));

describe('teams_get_notes input', () => {
  it('takes either a conversation or a Loop URL, and a tab only within a conversation', () => {
    expect(GetNotesInputSchema.safeParse({ conversationId: '19:a@thread.v2' }).success).toBe(true);
    expect(GetNotesInputSchema.safeParse({ conversationId: '19:a@thread.v2', tabId: 'notes:1' }).success).toBe(true);
    expect(GetNotesInputSchema.safeParse({ url: 'https://contoso.sharepoint.com/x.loop' }).success).toBe(true);
    expect(GetNotesInputSchema.safeParse({}).success).toBe(false);
    expect(GetNotesInputSchema.safeParse({ conversationId: '19:a@thread.v2', url: 'https://contoso.sharepoint.com/x.loop' }).success).toBe(false);
    expect(GetNotesInputSchema.safeParse({ url: 'https://contoso.sharepoint.com/x.loop', tabId: 'notes:1' }).success).toBe(false);
  });
});

describe('teams_get_notes with a Graph listing failure', () => {
  it('reports the Graph failure, not NOT_FOUND, for a tab only Graph could resolve', async () => {
    vi.mocked(listConversationTabs).mockResolvedValue(ok({
      conversationId: '19:a@thread.v2',
      conversationType: 'chat',
      tabs: [{ id: 'notes:1', name: 'Notes', kind: 'notes', readable: false, reason: 'unresolvable' }],
      tabsError: { code: ErrorCode.API_ERROR, message: 'HTTP 503: Service Unavailable', retryable: true },
    }));

    const result = await getNotesTool.handler({ conversationId: '19:a@thread.v2', tabId: 'graph-tab-1' });

    expect(result).toMatchObject({
      success: false,
      error: { code: ErrorCode.API_ERROR, retryable: true, message: expect.stringContaining('HTTP 503') },
    });
  });
});

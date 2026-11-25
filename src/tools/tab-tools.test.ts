import { describe, expect, it } from 'vitest';
import { GetNotesInputSchema } from './index.js';

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

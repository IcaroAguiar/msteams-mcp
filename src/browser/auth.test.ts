import { describe, expect, it, vi } from 'vitest';
import type { Page } from 'playwright';

vi.mock('./context.js', () => ({ saveSessionState: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ extractSubstrateToken: vi.fn() }));

import { getAuthStatus, isTeamsUrl } from './auth.js';

/** A page on `url` with no Teams DOM markers and `tokenMinutes` left on MSAL. */
function fakePage(url: string, tokenMinutes: number): Page {
  return {
    url: () => url,
    locator: () => ({ count: async () => 0 }),
    evaluate: async () => tokenMinutes,
  } as unknown as Page;
}

describe('isTeamsUrl', () => {
  it('covers every host Microsoft redirects Teams users to', () => {
    // New Teams settles on the bare host, with no path.
    for (const url of [
      'https://teams.microsoft.com/',
      'https://teams.cloud.microsoft/',
      'https://teams.chat.microsoft.com/',
      'https://teams.microsoft.us/',
      'https://dod.teams.microsoft.us/',
    ]) {
      expect(isTeamsUrl(url), url).toBe(true);
    }
    expect(isTeamsUrl('https://outlook.office.com/')).toBe(false);
  });

  it('matches on the host, whatever path the app has routed to', () => {
    expect(isTeamsUrl('https://teams.microsoft.com/v2/#/conversations/19:abc@thread.tacv2')).toBe(true);
  });
});

describe('getAuthStatus', () => {
  it('detects a completed login from the MSAL token alone', async () => {
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/', 55));

    expect(status.isAuthenticated).toBe(true);
    expect(status.isOnLoginPage).toBe(false);
  });

  it('stays unauthenticated on a Teams host with no token', async () => {
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/', -1));

    expect(status.isAuthenticated).toBe(false);
  });

  it('reports the login page', async () => {
    const status = await getAuthStatus(fakePage('https://login.microsoftonline.com/common/oauth2', 99));

    expect(status).toMatchObject({ isAuthenticated: false, isOnLoginPage: true });
  });
});

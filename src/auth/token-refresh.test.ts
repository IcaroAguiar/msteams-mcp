import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  getSessionIdentities: vi.fn(() => new Set<string>()),
}));

vi.mock('./session-store.js', () => ({
  clearTokenCache: vi.fn(),
  readSessionState: vi.fn(),
  writeSessionState: vi.fn(),
}));

vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn(), guardSessionWrites: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { refreshTokensViaBrowser, resetChromeRefreshCooldown } from './token-refresh.js';
import { extractSubstrateToken, getSessionIdentities } from './token-extractor.js';
import { readSessionState, writeSessionState } from './session-store.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { createBrowserContext, closeBrowser, guardSessionWrites } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { ErrorCode, createError } from '../types/errors.js';
import type { BrowserManager } from '../browser/context.js';
import type { SessionState } from './session-store.js';

describe('refreshTokensViaBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('uses HTTP refresh after the cached access token has expired', async () => {
    const refreshedExpiry = new Date(Date.now() + 60 * 60 * 1000);

    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce({ token: 'refreshed-access-token', expiry: refreshedExpiry });
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });

    const result = await refreshTokensViaBrowser();

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
    expect(result.ok).toBe(true);
  });

  describe('with the user\'s Chrome profile', () => {
    const expired = { ok: false as const, error: createError(ErrorCode.AUTH_EXPIRED, 'AADSTS700084: refresh token expired') };
    const page = { close: vi.fn().mockResolvedValue(undefined) };
    const manager = { kind: 'attached', page, context: {}, browser: {} } as unknown as BrowserManager;

    const stored: SessionState = { cookies: [], origins: [] };
    const ids = (...accounts: string[]) => new Set(accounts.map(id => `tenant/${id}`));
    const gateOf = () => vi.mocked(guardSessionWrites).mock.calls[0][1];

    beforeEach(() => {
      vi.stubEnv('TEAMS_EXISTING_CHROME_DIR', '/synthetic/chrome');
      resetChromeRefreshCooldown();
      vi.mocked(ensureAuthenticated).mockReset();
      vi.mocked(readSessionState).mockReturnValue(stored);
      vi.mocked(getSessionIdentities).mockReset().mockReturnValue(ids('A'));
      vi.mocked(refreshTokensViaHttp).mockResolvedValue(expired);
      vi.mocked(createBrowserContext).mockResolvedValue(manager);
      vi.mocked(closeBrowser).mockResolvedValue(undefined);
    });
    afterEach(() => vi.unstubAllEnvs());

    it('renews through a background Chrome tab once HTTP refresh can no longer', async () => {
      vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'fresh', expiry: new Date(Date.now() + 3600_000) });

      expect(await refreshTokensViaBrowser()).toEqual({ ok: true, value: undefined });
      expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
      expect(ensureAuthenticated).toHaveBeenCalledWith(page, manager.context, expect.any(Function), false, true);
      expect(closeBrowser).toHaveBeenCalledWith(manager, false);
      expect(page.close).toHaveBeenCalledOnce();
    });

    it('closes its tab when sign-in needs the user, then pauses further Chrome attempts', async () => {
      vi.mocked(ensureAuthenticated).mockRejectedValueOnce(new Error('Headless SSO failed: Login page detected'));

      const first = await refreshTokensViaBrowser();
      const second = await refreshTokensViaBrowser();

      expect(first).toMatchObject({ ok: false, error: { message: expect.stringContaining('Login page detected') } });
      expect(page.close).toHaveBeenCalledOnce();
      expect(second).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED, message: expect.stringContaining('paused') } });
      expect(createBrowserContext).toHaveBeenCalledOnce();

      resetChromeRefreshCooldown();
      await refreshTokensViaBrowser();
      expect(createBrowserContext).toHaveBeenCalledTimes(2);
    });


    /** Stored session holds `before`; the browser tries to save a session holding `after`. */
    const candidate: SessionState = { cookies: [], origins: [] };

    function browserSaves(before: Set<string>, after: Set<string>) {
      vi.mocked(getSessionIdentities).mockReset().mockImplementation(state => (state === stored ? before : after));
      vi.mocked(ensureAuthenticated).mockImplementationOnce(async () => gateOf()(candidate));
    }

    it('writes nothing when the browser is signed in as another account', async () => {
      browserSaves(ids('A'), ids('B'));

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_REQUIRED, message: expect.stringContaining('different Teams account') } });
      expect(writeSessionState).not.toHaveBeenCalled();
    });

    it('rejects a candidate that holds the stored account plus another', async () => {
      browserSaves(ids('A'), ids('A', 'B'));

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_REQUIRED } });
    });

    it('does not renew a stored session that holds several accounts, and never opens Chrome', async () => {
      browserSaves(ids('A', 'B'), ids('A', 'B'));

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_REQUIRED } });
      expect(createBrowserContext).not.toHaveBeenCalled();
    });

    it('does not renew without a stored account', async () => {
      browserSaves(ids(), ids('A'));

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_REQUIRED } });
      expect(createBrowserContext).not.toHaveBeenCalled();
    });

    it('keeps the stored session when cleanup fails after a rejected save', async () => {
      browserSaves(ids('A'), ids('B'));
      vi.mocked(closeBrowser).mockRejectedValue(new Error('disconnect failed'));

      expect((await refreshTokensViaBrowser()).ok).toBe(false);
      expect(writeSessionState).not.toHaveBeenCalled();
    });

    it('accepts a renewal that keeps the same single account', async () => {
      browserSaves(ids('A'), ids('A'));
      vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'fresh', expiry: new Date(Date.now() + 3600_000) });

      expect(await refreshTokensViaBrowser()).toEqual({ ok: true, value: undefined });
    });

    it('drops query strings and fragments from SSO URLs in the error', async () => {
      vi.mocked(ensureAuthenticated).mockRejectedValueOnce(new Error('Headless SSO failed: Unexpected page state: https://idp.example/path?code=SECRET&state=x#token=SECRET2'));

      const result = await refreshTokensViaBrowser();

      expect(JSON.stringify(result)).toContain('https://idp.example/path');
      expect(JSON.stringify(result)).not.toContain('SECRET');
    });

    it('does not ask for Chrome approval when HTTP refresh failed for a transient reason', async () => {
      vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: false, error: createError(ErrorCode.UNKNOWN, 'fetch failed', { retryable: true }) });

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { message: 'fetch failed' } });
      expect(createBrowserContext).not.toHaveBeenCalled();
    });
  });
});

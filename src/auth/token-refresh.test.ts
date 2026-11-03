import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
}));

vi.mock('./session-store.js', () => ({
  clearTokenCache: vi.fn(),
}));

vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { refreshTokensViaBrowser, resetChromeRefreshCooldown } from './token-refresh.js';
import { extractSubstrateToken } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { ErrorCode, createError } from '../types/errors.js';
import type { BrowserManager } from '../browser/context.js';

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

    beforeEach(() => {
      vi.stubEnv('TEAMS_EXISTING_CHROME_DIR', '/synthetic/chrome');
      resetChromeRefreshCooldown();
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

    it('does not ask for Chrome approval when HTTP refresh failed for a transient reason', async () => {
      vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: false, error: createError(ErrorCode.UNKNOWN, 'fetch failed', { retryable: true }) });

      expect(await refreshTokensViaBrowser()).toMatchObject({ ok: false, error: { message: 'fetch failed' } });
      expect(createBrowserContext).not.toHaveBeenCalled();
    });
  });
});

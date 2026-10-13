import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../auth/session-store.js', () => ({
  hasSessionState: vi.fn(() => true),
  isSessionLikelyExpired: vi.fn(() => false),
  clearSessionState: vi.fn(),
  clearTokenCache: vi.fn(),
}));

vi.mock('../auth/token-extractor.js', () => ({
  getSubstrateTokenStatus: vi.fn(),
  getMessageAuthStatus: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  extractSubstrateToken: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
}));

vi.mock('../auth/token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ resetAssignmentsAvailability: vi.fn() }));

import { loginTool } from './auth-tools.js';
import {
  getSubstrateTokenStatus,
  extractMessageAuth,
  extractCsaToken,
  extractSubstrateToken,
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';

/** A session holding every core credential except the CSA (favourites) token. */
function sessionMissingCsaToken(): void {
  vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
  vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3.6e6) });
  vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
  vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
  vi.mocked(extractCsaToken).mockReturnValue(null);
}

describe('teams_login core credential top-up', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
  });

  it('mints the missing CSA token when the session is otherwise valid', async () => {
    sessionMissingCsaToken();

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('skips the refresh when every core credential is present', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');

    await loginTool.handler({ forceNew: false });

    expect(refreshTokensViaHttp).not.toHaveBeenCalled();
  });

  it('tops up after a forced browser login leaves the CSA token behind', async () => {
    sessionMissingCsaToken();
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: { clearCookies: vi.fn() } } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await loginTool.handler({ forceNew: true });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
  });

  it('still reports success when the top-up fails', async () => {
    sessionMissingCsaToken();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
  });
});

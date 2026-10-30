import { randomBytes, randomUUID, webcrypto } from 'node:crypto';
import { afterEach, expect, it, vi } from 'vitest';
import { decodeMsalCache } from './msal-cache.js';
import { extractCsaToken, extractSubstrateToken, invalidateAccessToken } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { readSessionState, writeSessionState, type SessionState } from './session-store.js';

vi.mock('./session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('./session-store.js')>(),
  readSessionState: vi.fn(), writeSessionState: vi.fn(), clearTokenCache: vi.fn(),
}));

async function fixture() {
  const clientId = randomUUID();
  const key = randomBytes(32);
  const id = randomUUID();
  const secret = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ exp: Date.now() / 1000 + 3600 })).toString('base64url')}.${randomBytes(12).toString('base64url')}`;
  const refreshToken = randomBytes(32).toString('base64url');
  const baseKey = await webcrypto.subtle.importKey('raw', key, 'HKDF', false, ['deriveKey']);
  const encrypt = async (credentialType: string, token: string) => {
    const nonce = randomBytes(16);
    const derivedKey = await webcrypto.subtle.deriveKey({ name: 'HKDF', salt: nonce,
      hash: 'SHA-256', info: new TextEncoder().encode(clientId) }, baseKey,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
    const value = JSON.stringify({ credentialType, secret: token, clientId,
      target: 'SubstrateSearch', realm: randomUUID(), homeAccountId: randomUUID(), environment: 'login.windows.net' });
    const data = await webcrypto.subtle.encrypt({ name: 'AES-GCM', iv: new Uint8Array(12) }, derivedKey, new TextEncoder().encode(value));
    return { name: `msal.2.fixture-${credentialType}-${clientId}`,
      value: JSON.stringify({ id, nonce: nonce.toString('base64url'), data: Buffer.from(data).toString('base64url') }) };
  };
  const state: SessionState = {
    cookies: [{ name: 'msal.cache.encryption', domain: 'teams.cloud.microsoft', path: '/',
      value: encodeURIComponent(JSON.stringify({ id, key: key.toString('base64url') })) }],
    origins: [{ origin: 'https://teams.cloud.microsoft', localStorage: [
      { name: `msal.2.token.keys.${clientId}`, value: '{}' },
      await encrypt('AccessToken', secret), await encrypt('RefreshToken', refreshToken),
    ] }],
  };
  return { state, secret, refreshToken };
}

afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

it('extracts encrypted credentials without changing the source snapshot or plaintext entries', async () => {
  const { state, secret } = await fixture();
  const before = structuredClone(state);
  const decoded = decodeMsalCache(state);
  expect(extractSubstrateToken(state)?.token).toBe(secret);
  expect(state).toEqual(before);
  expect(decodeMsalCache(decoded)).toEqual(decoded);
  expect(decoded.origins[0].localStorage[0]).toEqual(before.origins[0].localStorage[0]);
});

it.each(['cookie-domain', 'cookie-parent', 'cookie-path', 'cookie-id', 'ciphertext', 'context', 'origin'])('rejects an encrypted cache with invalid %s', async failure => {
  const { state } = await fixture();
  if (failure === 'cookie-domain') state.cookies[0].domain = 'untrusted.example';
  if (failure === 'cookie-parent') state.cookies[0].domain = '.cloud.microsoft';
  if (failure === 'cookie-path') state.cookies[0].path = '/another-app';
  if (failure === 'cookie-id') state.cookies[0].value = encodeURIComponent(JSON.stringify({ id: randomUUID(), key: randomBytes(32).toString('base64url') }));
  if (failure === 'ciphertext') {
    const entry = JSON.parse(state.origins[0].localStorage[1].value);
    entry.data = randomBytes(64).toString('base64url');
    state.origins[0].localStorage[1].value = JSON.stringify(entry);
  }
  if (failure === 'context') state.origins[0].localStorage[0].name = `msal.2.token.keys.${randomUUID()}`;
  if (failure === 'origin') state.origins[0].origin = 'https://teams.cloud.microsoft.untrusted.example';
  expect(extractSubstrateToken(state)).toBeNull();
});

it('extracts CSA credentials and removes a rejected token from an encrypted snapshot', async () => {
  const { state, secret } = await fixture();
  state.origins[0].localStorage[1].name += '-chatsvcagg.teams.microsoft.com';
  expect(extractCsaToken(state)).toBe(secret);
  vi.mocked(readSessionState).mockReturnValue(state);
  invalidateAccessToken(secret);
  expect(vi.mocked(writeSessionState).mock.calls[0][0].origins[0].localStorage.map(entry => entry.name))
    .toEqual([state.origins[0].localStorage[0].name, state.origins[0].localStorage[2].name]);
});

it('renews an encrypted snapshot over HTTP with its decoded refresh token', async () => {
  const { state, secret, refreshToken } = await fixture();
  vi.mocked(readSessionState).mockReturnValue(state);
  const fetchMock = vi.fn().mockResolvedValue(new Response(JSON.stringify({ access_token: secret, expires_in: 3600, scope: 'Files.ReadWrite.All' }), { status: 200 }));
  vi.stubGlobal('fetch', fetchMock);
  expect(await refreshTokensViaHttp('graph')).toMatchObject({ ok: true });
  const body = new URLSearchParams(fetchMock.mock.calls[0][1].body);
  expect(body.get('refresh_token')).toBe(refreshToken);
  expect(writeSessionState).toHaveBeenCalledOnce();
});

import { createDecipheriv, hkdfSync } from 'node:crypto';
import { z } from 'zod';
import { TEAMS_HOSTS } from '../constants.js';
import type { SessionState } from './session-store.js';

const EncryptionCookie = z.object({ id: z.string(), key: z.string() });
const EncryptedEntry = z.object({ id: z.string(), nonce: z.string(), data: z.string() });

// Wire format: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/msal-browser-v5.6.3/lib/msal-browser/src/crypto/BrowserCrypto.ts
export function decodeMsalCache(state: SessionState): SessionState {
  return { ...state, origins: state.origins.map(origin => {
    if (!TEAMS_HOSTS.some(host => origin.origin === `https://${host}`)) return origin;
    const hostname = new URL(origin.origin).hostname;
    const cookie = state.cookies.find(candidate => candidate.name === 'msal.cache.encryption'
      && candidate.domain?.replace(/^\./, '') === hostname
      && (candidate.path === '/' || candidate.path === undefined));
    if (!cookie) return origin;
    try {
      const key = EncryptionCookie.parse(JSON.parse(decodeURIComponent(cookie.value)));
      const baseKey = Buffer.from(key.key, 'base64url');
      if (baseKey.length !== 32) return origin;
      const clientIds = origin.localStorage.flatMap(entry =>
        entry.name.match(/^msal(?:\.\d+)?\.token\.keys\.(.+)$/)?.slice(1) ?? []);
      return { ...origin, localStorage: origin.localStorage.map(entry => {
        if (!entry.name.startsWith('msal.')) return entry;
        try {
          const encrypted = EncryptedEntry.parse(JSON.parse(entry.value));
          if (encrypted.id !== key.id) return entry;
          const nonce = Buffer.from(encrypted.nonce, 'base64url');
          if (nonce.length !== 16) return entry;
          const context = clientIds.find(clientId => entry.name.includes(clientId)) ?? '';
          const derivedKey = hkdfSync('sha256', baseKey, nonce, context, 32);
          const data = Buffer.from(encrypted.data, 'base64url');
          const decipher = createDecipheriv('aes-256-gcm', Buffer.from(derivedKey), Buffer.alloc(12));
          decipher.setAuthTag(data.subarray(-16));
          const value = Buffer.concat([decipher.update(data.subarray(0, -16)), decipher.final()]).toString('utf8');
          JSON.parse(value);
          return { ...entry, value };
        } catch { return entry; }
      }) };
    } catch { return origin; }
  }) };
}

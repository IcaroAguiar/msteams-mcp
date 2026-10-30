/**
 * Encryption utilities for credential storage.
 * 
 * Uses a random key held by the operating system credential store.
 * The persistent browser profile remains sensitive and needs OS protection.
 */

import * as crypto from 'crypto';
import { Entry } from '@napi-rs/keyring';

/** Algorithm for encryption. */
const ALGORITHM = 'aes-256-gcm';

function credentialKey(create: boolean): Buffer {
  const entry = new Entry('msteams-mcp', 'session-encryption-v2', {
    linux: { store: 'secret-service' },
  });
  let secret = entry.getPassword();
  if (secret === null && create) {
    secret = crypto.randomBytes(32).toString('hex');
    entry.setPassword(secret);
  }
  if (secret === null || !/^[a-f0-9]{64}$/.test(secret)) {
    throw new Error('Session encryption key is missing or invalid in the OS credential store.');
  }
  return Buffer.from(secret, 'hex');
}

/** Encrypted data format. */
export interface EncryptedData {
  /** Initialisation vector (hex). */
  iv: string;
  /** Encrypted content (hex). */
  content: string;
  /** Authentication tag (hex). */
  tag: string;
  /** Version marker for format compatibility. */
  version: number;
}

/**
 * Encrypts a string value.
 */
export function encrypt(plaintext: string): EncryptedData {
  const key = credentialKey(true);
  const iv = crypto.randomBytes(16);
  
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  
  let encrypted = cipher.update(plaintext, 'utf8', 'hex');
  encrypted += cipher.final('hex');
  
  const tag = cipher.getAuthTag();
  
  return {
    iv: iv.toString('hex'),
    content: encrypted,
    tag: tag.toString('hex'),
    version: 2,
  };
}

/**
 * Decrypts encrypted data.
 * 
 * @throws Error if decryption fails (wrong machine, corrupted data, etc.)
 */
export function decrypt(data: EncryptedData): string {
  if (data.version !== 2) {
    throw new Error('Legacy or unsupported session encryption. Sign in again to capture a protected session.');
  }
  
  const key = credentialKey(false);
  const iv = Buffer.from(data.iv, 'hex');
  const tag = Buffer.from(data.tag, 'hex');
  
  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);
  
  let decrypted = decipher.update(data.content, 'hex', 'utf8');
  decrypted += decipher.final('utf8');
  
  return decrypted;
}

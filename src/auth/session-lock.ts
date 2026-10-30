import { lock } from 'proper-lockfile';
import { CONFIG_DIR, ensureConfigDir } from './session-store.js';

export async function withSessionLock<T>(operation: () => Promise<T>): Promise<T> {
  ensureConfigDir();
  const release = await lock(CONFIG_DIR, {
    realpath: false,
    lockfilePath: `${CONFIG_DIR}.operation.lock`,
    stale: 600_000,
    update: 1_000,
    retries: 0,
  });
  try {
    return await operation();
  } finally {
    await release();
  }
}

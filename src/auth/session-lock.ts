import { lock } from 'proper-lockfile';
import { CONFIG_DIR, ensureConfigDir } from './session-store.js';

// One HTTP server serves many clients from a single process, so operations queue
// here before taking the cross-process file lock (which never retries).
let queue: Promise<unknown> = Promise.resolve();

export async function withSessionLock<T>(operation: () => Promise<T>): Promise<T> {
  const run = queue.then(() => withFileLock(operation));
  queue = run.catch(() => undefined);
  return run;
}

async function withFileLock<T>(operation: () => Promise<T>): Promise<T> {
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

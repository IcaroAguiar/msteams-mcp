import { lock } from 'proper-lockfile';
import { CONFIG_DIR, ensureConfigDir } from './session-store.js';

// One HTTP server serves many clients from a single process, so operations queue
// here before taking the cross-process file lock (which never retries).
let queue: Promise<unknown> = Promise.resolve();
const QUEUE_WAIT_MS = 60_000;

export async function withSessionLock<T>(operation: () => Promise<T>, maxWaitMs = QUEUE_WAIT_MS): Promise<T> {
  const previous = queue;
  let done!: () => void;
  const turn = new Promise<void>((resolve) => { done = resolve; });
  queue = previous.then(() => turn);
  try {
    await waitFor(previous, maxWaitMs);
    return await withFileLock(operation);
  } finally {
    done();
  }
}

async function waitFor(previous: Promise<unknown>, maxWaitMs: number): Promise<void> {
  let timer: NodeJS.Timeout | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(Object.assign(new Error('Session busy'), { code: 'ELOCKED' })), maxWaitMs);
  });
  try {
    await Promise.race([previous, timeout]);
  } finally {
    clearTimeout(timer);
  }
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

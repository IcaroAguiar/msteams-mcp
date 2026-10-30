import { afterAll, beforeAll, expect, it, vi } from 'vitest';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { setTimeout } from 'node:timers/promises';

const directory = mkdtempSync(join(tmpdir(), 'teams-lock-'));
let withSessionLock: typeof import('./session-lock.js').withSessionLock;
beforeAll(async () => {
  vi.stubEnv('TEAMS_CONFIG_DIR', join(directory, 'session'));
  ({ withSessionLock } = await import('./session-lock.js'));
});
afterAll(() => { vi.unstubAllEnvs(); rmSync(directory, { recursive: true, force: true }); });

it('keeps another process out until the current operation finishes', async () => {
  const resultPath = join(directory, 'child-result');
  let child: ReturnType<typeof spawn> | undefined;
  let finished: Promise<unknown> | undefined;
  try {
    await withSessionLock(async () => {
      child = spawn(process.execPath, ['--import', 'tsx', '--input-type=module', '-e', `
        import { withSessionLock } from './src/auth/session-lock.ts';
        import { writeFileSync } from 'node:fs';
        process.stdout.write('waiting');
        await withSessionLock(async () => writeFileSync(process.argv[1], 'completed'));
      `, resultPath], { env: process.env, stdio: ['ignore', 'pipe', 'pipe'] });
      finished = once(child, 'exit');
      await once(child.stdout!, 'data');
      await setTimeout(100);
      expect(existsSync(resultPath)).toBe(false);
    });
    await finished;
    expect(existsSync(resultPath)).toBe(true);
  } finally {
    if (child && child.exitCode === null) child.kill();
  }
});

it('releases the lock after an operation fails', async () => {
  await expect(withSessionLock(async () => { throw new Error('fixture failure'); })).rejects.toThrow('fixture failure');
  await expect(withSessionLock(async () => 'next operation')).resolves.toBe('next operation');
});

import { afterAll, beforeAll, expect, it, vi } from 'vitest';
import { readFileSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { once } from 'node:events';

const directory = mkdtempSync(join(tmpdir(), 'teams-lock-'));
let withSessionLock: typeof import('./session-lock.js').withSessionLock;
beforeAll(async () => {
  vi.stubEnv('TEAMS_CONFIG_DIR', join(directory, 'session'));
  ({ withSessionLock } = await import('./session-lock.js'));
});
afterAll(() => { vi.unstubAllEnvs(); rmSync(directory, { recursive: true, force: true }); });

it('refuses another process while the current operation owns the session', async () => {
  const resultPath = join(directory, 'child-result');
  let child: ReturnType<typeof spawn> | undefined;
  let finished: Promise<unknown> | undefined;
  try {
    await withSessionLock(async () => {
      child = spawn(process.execPath, ['--import', 'tsx', '--input-type=module', '-e', `
        import { withSessionLock } from './src/auth/session-lock.ts';
        import { writeFileSync } from 'node:fs';
        process.stdout.write('waiting');
        try { await withSessionLock(async () => writeFileSync(process.argv[1], 'entered')); }
        catch (error) { writeFileSync(process.argv[1], error.code); }
      `, resultPath], { env: process.env, stdio: ['ignore', 'pipe', 'pipe'] });
      finished = once(child, 'exit');
      await once(child.stdout!, 'data');
      await finished;
      expect(readFileSync(resultPath, 'utf8')).toBe('ELOCKED');
    });

  } finally {
    if (child && child.exitCode === null) child.kill();
  }
});

it('releases the lock after an operation fails', async () => {
  await expect(withSessionLock(async () => { throw new Error('fixture failure'); })).rejects.toThrow('fixture failure');
  await expect(withSessionLock(async () => 'next operation')).resolves.toBe('next operation');
});

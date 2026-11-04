import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { BrowserContext } from 'playwright';

const mocks = vi.hoisted(() => ({ writeSessionState: vi.fn() }));
vi.mock('../auth/session-store.js', () => ({ ensureConfigDir: vi.fn(), CONFIG_DIR: '/synthetic', writeSessionState: mocks.writeSessionState }));
vi.mock('../utils/auth-guards.js', () => ({ clearRegionCache: vi.fn() }));
vi.mock('./existing-chrome.js', () => ({ captureTeamsPage: vi.fn(), connectExistingChrome: vi.fn(), existingChromeRequested: vi.fn() }));

import { guardSessionWrites, saveSessionState } from './context.js';

describe('guardSessionWrites', () => {
  const state = { cookies: [], origins: [] };
  const context = { storageState: vi.fn().mockResolvedValue(state) } as unknown as BrowserContext;

  beforeEach(() => vi.clearAllMocks());

  it('writes the captured session when the gate accepts it', async () => {
    guardSessionWrites(context, () => {});
    await saveSessionState(context);
    expect(mocks.writeSessionState).toHaveBeenCalledWith(state);
  });

  it('writes nothing when the gate rejects the captured session', async () => {
    guardSessionWrites(context, () => { throw new Error('other account'); });
    await expect(saveSessionState(context)).rejects.toThrow('other account');
    expect(mocks.writeSessionState).not.toHaveBeenCalled();
  });
});

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { randomUUID } from 'node:crypto';
import type { Browser, BrowserContext, Page } from 'playwright';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({ connect: vi.fn(), launch: vi.fn(), write: vi.fn() }));
vi.mock('playwright', () => ({ chromium: {
  connectOverCDP: mocks.connect, launchPersistentContext: mocks.launch,
} }));
vi.mock('../auth/session-store.js', () => ({
  CONFIG_DIR: '/unused', ensureConfigDir: vi.fn(), writeSessionState: mocks.write,
}));
vi.mock('../utils/auth-guards.js', () => ({ clearRegionCache: vi.fn() }));
import { closeBrowser, createBrowserContext, saveSessionState } from './context.js';

describe('existing Chrome profile', () => {
  let directory: string;
  beforeEach(() => {
    vi.clearAllMocks();
    directory = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-chrome-test-'));
    vi.stubEnv('TEAMS_EXISTING_CHROME_DIR', directory);
    fs.writeFileSync(path.join(directory, 'DevToolsActivePort'), '19222\n/devtools/browser/test-browser\n');
  });
  afterEach(() => {
    vi.unstubAllEnvs();
    fs.rmSync(directory, { recursive: true, force: true });
  });
  function chrome(url = 'https://teams.cloud.microsoft/') {
    const entry = { name: 'msal.test', value: randomUUID() };
    const context = {
      newPage: vi.fn(), cookies: vi.fn().mockResolvedValue([]),
      storageState: vi.fn(), close: vi.fn(),
    };
    const page = {
      url: () => url, context: () => context as unknown as BrowserContext,
      evaluate: vi.fn().mockResolvedValue({ origin: new URL(url).origin, localStorage: [entry] }),
      bringToFront: vi.fn().mockResolvedValue(undefined), close: vi.fn().mockResolvedValue(undefined),
    };
    context.newPage.mockResolvedValue(page as unknown as Page);
    const browser = { contexts: () => [context], close: vi.fn() };
    mocks.connect.mockResolvedValue(browser as unknown as Browser);
    return { context, page, browser, entry };
  }
  it('attaches locally and leaves tabs intact when login is incomplete', async () => {
    const { context, page, browser } = chrome();
    const manager = await createBrowserContext({ headless: false });
    await closeBrowser(manager, false);
    expect(mocks.connect).toHaveBeenCalledWith('ws://127.0.0.1:19222/devtools/browser/test-browser', { timeout: 60000 });
    expect(mocks.launch).not.toHaveBeenCalled();
    expect(page.bringToFront).toHaveBeenCalledOnce();
    expect(page.close).not.toHaveBeenCalled();
    expect(context.close).not.toHaveBeenCalled();
    expect(browser.close).toHaveBeenCalledOnce();
    expect(mocks.write).not.toHaveBeenCalled();
  });
  it('captures only the Teams tab and closes only that new tab after success', async () => {
    const { context, page, browser, entry } = chrome();
    const manager = await createBrowserContext({ headless: false });
    await closeBrowser(manager, true);
    expect(context.storageState).not.toHaveBeenCalled();
    expect(context.cookies).toHaveBeenCalledWith('https://teams.cloud.microsoft');
    expect(mocks.write).toHaveBeenCalledWith({ cookies: [], origins: [
      { origin: 'https://teams.cloud.microsoft', localStorage: [entry] },
    ] });
    expect(page.close).toHaveBeenCalledOnce();
    expect(context.close).not.toHaveBeenCalled();
    expect(browser.close).toHaveBeenCalledOnce();
  });
  it('rejects capture of another site and leaves its tab open', async () => {
    const { context, page, browser } = chrome('https://login.microsoftonline.com/');
    const manager = await createBrowserContext({ headless: false });
    await expect(closeBrowser(manager, true)).rejects.toThrow('no session was captured');
    expect(page.evaluate).not.toHaveBeenCalled();
    expect(context.cookies).not.toHaveBeenCalled();
    expect(context.storageState).not.toHaveBeenCalled();
    expect(page.close).not.toHaveBeenCalled();
    expect(mocks.write).not.toHaveBeenCalled();
    expect(browser.close).toHaveBeenCalledOnce();
  });
  it('uses the same scoped capture while authentication is polling', async () => {
    const { entry, context } = chrome();
    const manager = await createBrowserContext({ headless: false });
    await saveSessionState(manager.context);
    expect(mocks.write.mock.calls[0][0].origins).toEqual([
      { origin: 'https://teams.cloud.microsoft', localStorage: [entry] },
    ]);
    expect(context.storageState).not.toHaveBeenCalled();
    await closeBrowser(manager, false);
  });
  it('rejects capture if the page navigated away before evaluation', async () => {
    const { page, context } = chrome();
    page.evaluate.mockResolvedValue(null);
    const manager = await createBrowserContext({ headless: false });
    await expect(closeBrowser(manager, true)).rejects.toThrow('left Teams');
    expect(context.cookies).not.toHaveBeenCalled();
    expect(mocks.write).not.toHaveBeenCalled();
    expect(page.close).not.toHaveBeenCalled();
  });
  it.each(['70000\n/devtools/browser/test', '9222\nwss://elsewhere.example/browser', 'bad'])('rejects invalid endpoint data: %s', async contents => {
    fs.writeFileSync(path.join(directory, 'DevToolsActivePort'), contents);
    await expect(createBrowserContext({ headless: false })).rejects.toThrow('enable remote debugging');
    expect(mocks.connect).not.toHaveBeenCalled();
    expect(mocks.launch).not.toHaveBeenCalled();
  });
  function backgroundTarget(context: object, pages: Record<string, unknown>) {
    const ids = new Map(Object.entries(pages).map(([id, p]) => [p, id]));
    Object.assign(context, {
      newCDPSession: vi.fn(async (p: unknown) => ({
        send: vi.fn().mockResolvedValue({ targetInfo: { targetId: ids.get(p) } }),
        detach: vi.fn().mockResolvedValue(undefined),
      })),
      waitForEvent: vi.fn(async (_event: string, { predicate }: { predicate: (p: unknown) => Promise<boolean> }) => {
        for (const p of Object.values(pages)) if (await predicate(p)) return p;
        throw new Error('Timeout 15000ms exceeded');
      }),
    });
  }
  it('renews automatically in a background tab that does not take focus', async () => {
    const { context, page, browser } = chrome();
    const usersTab = { close: vi.fn(), bringToFront: vi.fn() };
    const session = { send: vi.fn().mockResolvedValue({ targetId: 'ours' }), detach: vi.fn().mockResolvedValue(undefined) };
    Object.assign(browser, { newBrowserCDPSession: vi.fn().mockResolvedValue(session) });
    backgroundTarget(context, { theirs: usersTab, ours: page });

    const manager = await createBrowserContext({ headless: true });

    expect(manager.page).toBe(page);
    expect(usersTab.close).not.toHaveBeenCalled();
    expect(session.send).toHaveBeenCalledWith('Target.createTarget', { url: 'about:blank', background: true });
    expect(session.detach).toHaveBeenCalledOnce();
    expect(context.newPage).not.toHaveBeenCalled();
    expect(page.bringToFront).not.toHaveBeenCalled();
    expect(mocks.launch).not.toHaveBeenCalled();
  });
  it('closes the background tab it created when the page never appears', async () => {
    const { context, browser } = chrome();
    const session = { send: vi.fn().mockResolvedValueOnce({ targetId: 'orphan' }).mockResolvedValue({}), detach: vi.fn().mockResolvedValue(undefined) };
    Object.assign(browser, { newBrowserCDPSession: vi.fn().mockResolvedValue(session) });
    backgroundTarget(context, {});

    await expect(createBrowserContext({ headless: true })).rejects.toThrow('Timeout');

    expect(session.send).toHaveBeenCalledWith('Target.closeTarget', { targetId: 'orphan' });
    expect(browser.close).toHaveBeenCalledOnce();
  });
  it('fails cleanly when Chrome refuses to create the tab', async () => {
    const { context, browser } = chrome();
    const session = { send: vi.fn().mockRejectedValue(new Error('Target.createTarget refused')), detach: vi.fn().mockResolvedValue(undefined) };
    Object.assign(browser, { newBrowserCDPSession: vi.fn().mockResolvedValue(session) });
    Object.assign(context, { waitForEvent: vi.fn(() => new Promise(() => {})) });
    const unhandled = vi.fn();
    process.on('unhandledRejection', unhandled);
    try {
      await expect(createBrowserContext({ headless: true })).rejects.toThrow('refused');
      await new Promise(resolve => setTimeout(resolve, 10));
      expect(unhandled).not.toHaveBeenCalled();
      expect(session.detach).toHaveBeenCalledOnce();
      expect(browser.close).toHaveBeenCalledOnce();
    } finally {
      process.off('unhandledRejection', unhandled);
    }
  });
  it('keeps connection errors free of the debugging endpoint', async () => {
    mocks.connect.mockRejectedValue(new Error('ws://127.0.0.1:19222/devtools/browser/private-id'));
    const message = await createBrowserContext({ headless: false }).then(() => '', (error: Error) => error.message);
    expect(message).toContain('Approve its debugging connection prompt');
    expect(message).not.toMatch(/private-id|19222|ws:\/\//);
  });
  it('rejects a relative profile directory without echoing it', async () => {
    vi.stubEnv('TEAMS_EXISTING_CHROME_DIR', 'private-relative-directory');
    const message = await createBrowserContext({ headless: false }).then(() => '', (error: Error) => error.message);
    expect(message).toContain('must be an absolute path');
    expect(message).not.toContain('private-relative-directory');
    expect(mocks.connect).not.toHaveBeenCalled();
  });
  it('cleans up its blank tab if foreground activation fails', async () => {
    const { page, browser } = chrome();
    page.bringToFront.mockRejectedValueOnce(new Error('activation failed'));
    await expect(createBrowserContext({ headless: false })).rejects.toThrow('activation failed');
    expect(page.close).toHaveBeenCalledOnce();
    expect(browser.close).toHaveBeenCalledOnce();
  });
  it('keeps a captured login successful when the user closed its tab', async () => {
    const { page, browser } = chrome();
    const manager = await createBrowserContext({ headless: false });
    page.close.mockRejectedValueOnce(new Error('tab already closed'));
    await expect(closeBrowser(manager, true)).resolves.toBeUndefined();
    expect(mocks.write).toHaveBeenCalledOnce();
    expect(browser.close).toHaveBeenCalledOnce();
  });
});

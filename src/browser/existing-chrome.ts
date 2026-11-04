import * as fs from 'node:fs';
import * as path from 'node:path';
import { chromium, type Browser, type BrowserContext, type Page } from 'playwright';
import { z } from 'zod';
import { TEAMS_HOSTS } from '../constants.js';
import type { SessionState } from '../auth/session-store.js';
import type { BrowserManager } from './context.js';

export function existingChromeRequested(): boolean {
  return Boolean(process.env.TEAMS_EXISTING_CHROME_DIR);
}

export async function connectExistingChrome(
  { background = false }: { background?: boolean } = {},
): Promise<Extract<BrowserManager, { kind: 'attached' }>> {
  const directory = process.env.TEAMS_EXISTING_CHROME_DIR;
  if (!directory) throw new Error('TEAMS_EXISTING_CHROME_DIR is not configured.');
  if (!path.isAbsolute(directory)) throw new Error('TEAMS_EXISTING_CHROME_DIR must be an absolute path.');
  let endpoint: string;
  try {
    const [port, route] = fs.readFileSync(path.join(directory, 'DevToolsActivePort'), 'utf8').trim().split(/\r?\n/);
    const parsed = z.object({
      port: z.coerce.number().int().min(1).max(65535),
      route: z.string().regex(/^\/devtools\/browser\/[a-zA-Z0-9-]+$/),
    }).parse({ port, route });
    endpoint = `ws://127.0.0.1:${parsed.port}${parsed.route}`;
  } catch {
    throw new Error('Open Chrome 144+ and enable remote debugging at chrome://inspect/#remote-debugging, then retry teams_login.');
  }

  let browser: Browser;
  try {
    browser = await chromium.connectOverCDP(endpoint, { timeout: 60000 });
  } catch {
    throw new Error('Could not connect to the open Chrome. Approve its debugging connection prompt and retry teams_login.');
  }
  try {
    const context = browser.contexts()[0];
    if (!context) throw new Error('Chrome did not expose its existing profile.');
    if (background) return { kind: 'attached', browser, context, page: await openBackgroundTab(browser, context) };
    const page = await context.newPage();
    try {
      await page.bringToFront();
    } catch (error) {
      await page.close().catch(() => {});
      throw error;
    }
    return { kind: 'attached', browser, context, page };
  } catch (error) {
    await browser.close();
    throw error;
  }
}

async function openBackgroundTab(browser: Browser, context: BrowserContext): Promise<Page> {
  const session = await browser.newBrowserCDPSession();
  const created = session.send('Target.createTarget', { url: 'about:blank', background: true });
  try {
    // Another tab may open at the same moment; only the target we created may be driven and closed.
    const [page] = await Promise.all([
      context.waitForEvent('page', {
        timeout: 15000,
        predicate: async candidate => (await targetIdOf(context, candidate)) === (await created).targetId,
      }),
      created,
    ]);
    return page;
  } catch (error) {
    const targetId = await created.then(target => target.targetId, () => undefined);
    if (targetId) await session.send('Target.closeTarget', { targetId }).catch(() => {});
    throw error;
  } finally {
    await session.detach().catch(() => {});
  }
}

async function targetIdOf(context: BrowserContext, page: Page): Promise<string | undefined> {
  const session = await context.newCDPSession(page).catch(() => undefined);
  if (!session) return undefined;
  try {
    return (await session.send('Target.getTargetInfo')).targetInfo.targetId;
  } catch {
    return undefined;
  } finally {
    await session.detach().catch(() => {});
  }
}

export async function captureTeamsPage(page: Page): Promise<SessionState> {
  const origin = new URL(page.url()).origin;
  if (!TEAMS_HOSTS.some(host => origin === `https://${host}`)) {
    throw new Error('The login tab has not reached Teams; no session was captured.');
  }
  const state = await page.evaluate(hosts => {
    if (location.protocol !== 'https:' || !hosts.includes(location.hostname)) return null;
    return {
      origin: location.origin,
      localStorage: Object.entries(window.localStorage).map(([name, value]) => ({ name, value })),
    };
  }, [...TEAMS_HOSTS]);
  if (!state) throw new Error('The login tab left Teams; no session was captured.');
  return { cookies: await page.context().cookies(state.origin), origins: [state] };
}

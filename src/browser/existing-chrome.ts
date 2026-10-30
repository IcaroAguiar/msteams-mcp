import * as fs from 'node:fs';
import * as path from 'node:path';
import { chromium, type Browser, type Page } from 'playwright';
import { z } from 'zod';
import { TEAMS_HOSTS } from '../constants.js';
import type { SessionState } from '../auth/session-store.js';
import type { BrowserManager } from './context.js';

export function existingChromeRequested(): boolean {
  return Boolean(process.env.TEAMS_EXISTING_CHROME_DIR);
}

export async function connectExistingChrome(): Promise<Extract<BrowserManager, { kind: 'attached' }>> {
  const directory = process.env.TEAMS_EXISTING_CHROME_DIR;
  if (!directory) throw new Error('TEAMS_EXISTING_CHROME_DIR is not configured.');
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

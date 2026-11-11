import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AddressInfo } from 'node:net';
import type { Server } from 'node:http';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

const directory = mkdtempSync(join(tmpdir(), 'teams-http-'));
const token = 'test-token';
let http: Server;
let base: string;

beforeAll(async () => {
  vi.stubEnv('TEAMS_CONFIG_DIR', join(directory, 'session'));
  const { runHttpServer } = await import('./http.js');
  http = await runHttpServer({ host: '127.0.0.1', port: 0, token });
  base = `http://127.0.0.1:${(http.address() as AddressInfo).port}`;
});
afterAll(() => {
  http?.close();
  vi.unstubAllEnvs();
  rmSync(directory, { recursive: true, force: true });
});

function connect(auth = token) {
  const client = new Client({ name: 'test', version: '1.0.0' });
  const transport = new StreamableHTTPClientTransport(new URL(`${base}/mcp`), {
    requestInit: { headers: { Authorization: `Bearer ${auth}` } },
  });
  return { client, ready: client.connect(transport) };
}

describe('http transport', () => {
  it('serves several clients from one process', async () => {
    const a = connect();
    const b = connect();
    await Promise.all([a.ready, b.ready]);
    const [toolsA, toolsB] = await Promise.all([a.client.listTools(), b.client.listTools()]);
    expect(toolsA.tools.length).toBeGreaterThan(0);
    expect(toolsB.tools.map((t) => t.name)).toEqual(toolsA.tools.map((t) => t.name));
    const health = await (await fetch(`${base}/health`)).json();
    expect(health.sessions).toBe(2);
    await Promise.all([a.client.close(), b.client.close()]);
  });

  it('rejects a wrong bearer token', async () => {
    const { ready } = connect('wrong');
    await expect(ready).rejects.toThrow();
  });

  it('refuses a non-loopback bind without a token', async () => {
    const { runHttpServer } = await import('./http.js');
    await expect(runHttpServer({ host: '0.0.0.0', port: 0 })).rejects.toThrow(/TEAMS_MCP_TOKEN/);
  });
});

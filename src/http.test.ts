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
    const health = await (await fetch(`${base}/health`, { headers: { Authorization: `Bearer ${token}` } })).json();
    expect(health.sessions).toBe(2);
    await Promise.all([a.client.close(), b.client.close()]);
  });

  it('rejects a wrong bearer token', async () => {
    const { ready } = connect('wrong');
    await expect(ready).rejects.toThrow();
  });

  it('rejects a foreign Origin or Host (DNS rebinding)', async () => {
    const headers = { Authorization: `Bearer ${token}` };
    expect((await fetch(`${base}/health`, { headers: { ...headers, Origin: 'https://evil.example' } })).status).toBe(403);
    const { request } = await import('node:http');
    const status = await new Promise<number>((resolve, reject) => {
      const req = request(`${base}/health`, { headers: { ...headers, Host: 'evil.example' } }, (res) => resolve(res.statusCode ?? 0));
      req.on('error', reject);
      req.end();
    });
    expect(status).toBe(403);
  });

  it('answers malformed JSON with 400', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'content-type': 'application/json' },
      body: '{not json',
    });
    expect(res.status).toBe(400);
  });

  it('expires an idle session but keeps one with a request in flight', async () => {
    const { runHttpServer } = await import('./http.js');
    const short = await runHttpServer({ host: '127.0.0.1', port: 0, token, idleMs: 100 });
    const url = new URL(`http://127.0.0.1:${(short.address() as AddressInfo).port}/mcp`);
    try {
      const headers = { Authorization: `Bearer ${token}`, 'content-type': 'application/json', accept: 'application/json, text/event-stream' };
      const init = await fetch(url, { method: 'POST', headers, body: JSON.stringify({
        jsonrpc: '2.0', id: 1, method: 'initialize',
        params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'raw', version: '1' } },
      }) });
      const sessionId = init.headers.get('mcp-session-id')!;
      await init.text();
      await new Promise((r) => setTimeout(r, 400));
      const after = await fetch(url, { method: 'POST', headers: { ...headers, 'mcp-session-id': sessionId },
        body: JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list' }) });
      expect(after.status).toBe(404);
      const { withSessionLock } = await import('./auth/session-lock.js');
      const init2 = await fetch(url, { method: 'POST', headers, body: JSON.stringify({
        jsonrpc: '2.0', id: 1, method: 'initialize',
        params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'raw-busy', version: '1' } },
      }) });
      const busyId = init2.headers.get('mcp-session-id')!;
      await init2.text();
      const sessionHeaders = { ...headers, 'mcp-session-id': busyId, 'mcp-protocol-version': '2025-06-18' };
      let release!: () => void;
      const holder = withSessionLock(() => new Promise<void>((r) => { release = r; }));
      const pending = fetch(url, { method: 'POST', headers: sessionHeaders,
        body: JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'teams_status', arguments: {} } }) })
        .then((r) => r.text());
      await new Promise((r) => setTimeout(r, 400));
      release();
      await holder;
      expect(await pending).toContain('"id":2');
      const still = await fetch(url, { method: 'POST', headers: sessionHeaders,
        body: JSON.stringify({ jsonrpc: '2.0', id: 3, method: 'tools/list' }) });
      expect(still.status).toBe(200);
    } finally {
      short.close();
    }
  });

  it('refuses a non-loopback bind without a token', async () => {
    const { runHttpServer } = await import('./http.js');
    await expect(runHttpServer({ host: '0.0.0.0', port: 0 })).rejects.toThrow(/TEAMS_MCP_TOKEN/);
  });
});

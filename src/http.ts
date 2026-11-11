/**
 * Streamable HTTP transport: one long-lived process shared by every local MCP client,
 * instead of one stdio process per agent session.
 */
import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { randomUUID, timingSafeEqual } from 'node:crypto';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import { TeamsServer } from './server.js';
import * as log from './utils/logger.js';

export interface HttpOptions {
  host: string;
  port: number;
  token?: string;
  /** Sessions with no request for this long are closed. */
  idleMs?: number;
}

interface Session {
  transport: StreamableHTTPServerTransport;
  lastSeen: number;
}

const LOOPBACK = new Set(['127.0.0.1', '::1', 'localhost']);
const MAX_BODY_BYTES = 4 * 1024 * 1024;

export async function runHttpServer(options: HttpOptions) {
  if (!LOOPBACK.has(options.host) && !options.token) {
    throw new Error('Refusing to listen beyond loopback without TEAMS_MCP_TOKEN');
  }
  const idleMs = options.idleMs ?? 30 * 60_000;
  const sessions = new Map<string, Session>();
  const teams = new TeamsServer();

  const sweep = setInterval(() => {
    const cutoff = Date.now() - idleMs;
    for (const session of sessions.values()) {
      if (session.lastSeen < cutoff) void session.transport.close();
    }
  }, 60_000);
  sweep.unref();

  const http = createHttpServer((req, res) => {
    handle(req, res).catch((error) => {
      log.error('http', `request failed: ${error instanceof Error ? error.message : String(error)}`);
      if (!res.headersSent) sendError(res, 500, 'Internal server error');
    });
  });

  async function handle(req: IncomingMessage, res: ServerResponse) {
    const url = new URL(req.url ?? '/', 'http://localhost');
    if (url.pathname === '/health') {
      res.writeHead(200, { 'content-type': 'application/json' }).end(JSON.stringify({ ok: true, sessions: sessions.size }));
      return;
    }
    if (url.pathname !== '/mcp') return sendError(res, 404, 'Not found');
    if (!authorized(req, options.token)) return sendError(res, 401, 'Unauthorized');

    const sessionId = header(req, 'mcp-session-id');
    const existing = sessionId ? sessions.get(sessionId) : undefined;
    if (existing) {
      existing.lastSeen = Date.now();
      const body = req.method === 'POST' ? await readJson(req) : undefined;
      await existing.transport.handleRequest(req, res, body);
      return;
    }
    if (sessionId) return sendError(res, 404, 'Session not found');
    if (req.method !== 'POST') return sendError(res, 400, 'Missing mcp-session-id');

    const body = await readJson(req);
    if (!isInitializeRequest(body)) return sendError(res, 400, 'First request must be initialize');

    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (id) => { sessions.set(id, { transport, lastSeen: Date.now() }); },
    });
    transport.onclose = () => { if (transport.sessionId) sessions.delete(transport.sessionId); };
    const server = await teams.createServer();
    await server.connect(transport);
    await transport.handleRequest(req, res, body);
  }

  await new Promise<void>((resolve, reject) => {
    http.once('error', reject);
    http.listen(options.port, options.host, () => resolve());
  });
  log.info('http', `listening on http://${options.host}:${options.port}/mcp`);

  const shutdown = () => {
    clearInterval(sweep);
    for (const { transport } of sessions.values()) void transport.close();
    http.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 2_000).unref();
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
  return http;
}

function header(req: IncomingMessage, name: string): string | undefined {
  const value = req.headers[name];
  return Array.isArray(value) ? value[0] : value;
}

function authorized(req: IncomingMessage, token: string | undefined): boolean {
  if (!token) return true;
  const given = Buffer.from(header(req, 'authorization') ?? '');
  const expected = Buffer.from(`Bearer ${token}`);
  return given.length === expected.length && timingSafeEqual(given, expected);
}

async function readJson(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > MAX_BODY_BYTES) throw new Error('Request body too large');
    chunks.push(chunk as Buffer);
  }
  const text = Buffer.concat(chunks).toString('utf8');
  return text ? JSON.parse(text) : undefined;
}

function sendError(res: ServerResponse, status: number, message: string) {
  res.writeHead(status, { 'content-type': 'application/json' })
    .end(JSON.stringify({ jsonrpc: '2.0', error: { code: -32000, message }, id: null }));
}

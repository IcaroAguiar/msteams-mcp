#!/usr/bin/env node
/**
 * Teams MCP Server entry point.
 *
 * This MCP server enables AI assistants to interact with Microsoft Teams
 * via direct API calls — searching messages, sending replies, managing
 * favourites, calendar, files, and more. The browser is only used for
 * initial authentication.
 */

import { runServer } from './server.js';
import { runHttpServer } from './http.js';

const httpPort = process.env.TEAMS_MCP_HTTP_PORT;
const start = process.argv.includes('--http') || httpPort
  ? runHttpServer({
      host: process.env.TEAMS_MCP_HTTP_HOST ?? '127.0.0.1',
      port: Number(httpPort ?? 47830),
      token: process.env.TEAMS_MCP_TOKEN || undefined,
    })
  : runServer();

start.catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});

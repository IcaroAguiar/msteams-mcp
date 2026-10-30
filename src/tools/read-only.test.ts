import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import type { Server } from '@modelcontextprotocol/sdk/server/index.js';

let directory: string;
let server: Server;
const client = new Client({ name: 'read-only-test', version: '1' });
beforeAll(async () => {
  directory = mkdtempSync(join(tmpdir(), 'teams-policy-'));
  vi.stubEnv('TEAMS_CONFIG_DIR', join(directory, 'session'));
  vi.stubEnv('TEAMS_READ_ONLY', '1');
  const { createServer } = await import('../server.js');
  server = await createServer();
  const [caller, receiver] = InMemoryTransport.createLinkedPair();
  await server.connect(receiver);
  await client.connect(caller);
});
afterAll(async () => {
  await client.close();
  await server.close();
  vi.unstubAllEnvs();
  rmSync(directory, { recursive: true, force: true });
});

describe('read-only MCP access', () => {
  it('advertises only the explicit chat-reading and login tools', async () => {
    const result = await client.listTools();
    expect(result.tools.map(tool => tool.name).sort()).toEqual([
      'teams_status', 'teams_login', 'teams_get_me', 'teams_search_people',
      'teams_get_person', 'teams_list_chats', 'teams_search', 'teams_get_thread',
      'teams_get_message', 'teams_get_chat', 'teams_get_unread',
      'teams_find_channel', 'teams_get_favorites', 'teams_get_activity',
    ].sort());
  });

  it.each([
    'teams_send_message', 'teams_delete_message', 'teams_edit_message',
    'teams_create_group_chat', 'teams_mark_read', 'teams_add_reaction',
    'teams_search_email', 'teams_download_file', 'teams_submission_action',
  ])('rejects a direct call to hidden tool %s', async name => {
    const result = await client.callTool({ name, arguments: {} });
    expect(result.isError).toBe(true);
    expect(result.content).toEqual([expect.objectContaining({
      text: expect.stringContaining('ACCESS_DENIED'),
    })]);
  });

  it.each([
    { name: 'teams_get_thread', arguments: { conversationId: 'fixture', markRead: true } },
    { name: 'teams_login', arguments: { forceNew: true } },
  ])('blocks mutating options on $name', async request => {
    const result = await client.callTool(request);
    expect(result.isError).toBe(true);
    expect(result.content).toEqual([expect.objectContaining({ text: expect.stringContaining('ACCESS_DENIED') })]);
  });

  it('reports an unauthenticated status without starting login', async () => {
    const result = await client.callTool({ name: 'teams_status', arguments: {} });
    expect(result.isError).not.toBe(true);
    expect(result.content).toEqual([expect.objectContaining({ text: expect.stringContaining('"exists": false') })]);
  });

  it('returns a structured busy response while another operation owns the session', async () => {
    const { withSessionLock } = await import('../auth/session-lock.js');
    await withSessionLock(async () => {
      const result = await client.callTool({ name: 'teams_status', arguments: {} });
      expect(result.isError).toBe(true);
      expect(result.content).toEqual([expect.objectContaining({ text: expect.stringContaining('RATE_LIMITED') })]);
    });
  });
});

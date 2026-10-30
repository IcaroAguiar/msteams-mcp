/**
 * Tool registry - centralised tool management.
 * 
 * All tools are registered here and can be looked up by name.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { z } from 'zod';
import type { ToolResult } from './index.js';
import { ErrorCode, createError } from '../types/errors.js';

/** Type-erased tool entry for the registry — avoids generic variance issues. */
interface RegistryEntry {
  definition: Tool;
  schema: z.ZodTypeAny;
  handler(input: z.output<z.ZodTypeAny>): Promise<ToolResult>;
}

import { searchTools } from './search-tools.js';
import { messageTools } from './message-tools.js';
import { peopleTools } from './people-tools.js';
import { authTools } from './auth-tools.js';
import { meetingTools } from './meeting-tools.js';
import { fileTools } from './file-tools.js';
import { tagTools } from './tag-tools.js';
import { assignmentTools } from './assignment-tools.js';

// ─────────────────────────────────────────────────────────────────────────────
// Registry
// ─────────────────────────────────────────────────────────────────────────────

/** All registered tools (cast to base type for registry). */
const allTools: RegistryEntry[] = [
  ...searchTools,
  ...messageTools,
  ...peopleTools,
  ...authTools,
  ...meetingTools,
  ...fileTools,
  ...tagTools,
  ...assignmentTools,
];

/**
 * Tool names grouped by category, for `msteams help`. Derived from the same
 * arrays as the registry, so a new tool appears in help automatically.
 */
export const toolCategories: Array<{ name: string; tools: string[] }> = [
  { name: 'Account', tools: [...authTools, ...peopleTools].map(t => t.definition.name) },
  { name: 'Search', tools: searchTools.map(t => t.definition.name) },
  { name: 'Messages', tools: messageTools.map(t => t.definition.name) },
  { name: 'Meetings', tools: meetingTools.map(t => t.definition.name) },
  { name: 'Files', tools: [...fileTools, ...tagTools].map(t => t.definition.name) },
  { name: 'Assignments', tools: assignmentTools.map(t => t.definition.name) },
];

/** Lookup map for tools by name. */
const toolsByName = new Map<string, RegistryEntry>(
  allTools.map(tool => [tool.definition.name, tool])
);

const readTools = new Set([
  'teams_status', 'teams_login', 'teams_get_me', 'teams_search_people',
  'teams_get_person', 'teams_list_chats', 'teams_search', 'teams_get_thread',
  'teams_get_message', 'teams_get_chat', 'teams_get_unread',
  'teams_find_channel', 'teams_get_favorites', 'teams_get_activity',
]);

function allowed(name: string): boolean {
  return process.env.TEAMS_READ_ONLY !== '1' || readTools.has(name);
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets all tool definitions for MCP ListTools.
 */
export function getToolDefinitions(): Tool[] {
  return allTools.filter(tool => allowed(tool.definition.name)).map(tool => tool.definition);
}

/**
 * Gets a tool by name.
 */
export function getTool(name: string): RegistryEntry | undefined {
  return allowed(name) ? toolsByName.get(name) : undefined;
}

/**
 * Invokes a tool by name with the given arguments.
 */
export async function invokeTool(name: string, args: unknown): Promise<ToolResult> {
  if (!allowed(name) || (process.env.TEAMS_READ_ONLY === '1' && args !== null && typeof args === 'object'
    && (('markRead' in args && args.markRead === true) || ('forceNew' in args && args.forceNew === true)))) {
    return { success: false, error: createError(ErrorCode.ACCESS_DENIED,
      'This operation is disabled by the read-only Teams policy.', { retryable: false }) };
  }
  const tool = toolsByName.get(name);
  
  if (!tool) {
    return {
      success: false,
      error: createError(ErrorCode.INVALID_INPUT, `Unknown tool: ${name}`, { retryable: false }),
    };
  }

  // Validate input
  const parseResult = tool.schema.safeParse(args);
  if (!parseResult.success) {
    return {
      success: false,
      error: createError(ErrorCode.INVALID_INPUT, `Invalid input: ${parseResult.error.message}`, { retryable: false }),
    };
  }

  // Invoke handler
  return tool.handler(parseResult.data);
}

/**
 * Checks if a tool exists.
 */
export function hasTool(name: string): boolean {
  return getTool(name) !== undefined;
}

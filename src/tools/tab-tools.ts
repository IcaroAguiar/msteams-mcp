/** Conversation tab tools (read-only): list tabs and read Notes/Loop pages as text. */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { listConversationTabs, readLoopPage, type ConversationTab } from '../api/tabs-api.js';
import { ErrorCode, createError } from '../types/errors.js';
import { NOTES_SCAN_MAX_PAGES, NOTES_SCAN_PAGE_SIZE } from '../constants.js';

export const GetTabsInputSchema = z.object({
  conversationId: z.string().min(1),
});

export const GetNotesInputSchema = z.object({
  conversationId: z.string().min(1).optional(),
  tabId: z.string().min(1).optional(),
  url: z.string().url().optional(),
}).refine(input => !!input.conversationId !== !!input.url, {
  message: 'Provide either conversationId or url (not both)',
}).refine(input => !input.tabId || !!input.conversationId, {
  message: 'tabId requires conversationId',
});

const SEARCH_DEPTH = NOTES_SCAN_MAX_PAGES * NOTES_SCAN_PAGE_SIZE;

const getTabsToolDefinition: Tool = {
  name: 'teams_get_tabs',
  description: `List the tabs of a Teams chat or channel: name, kind, app, content URL and Teams link. kind is one of notes (a chat's Notes page), loop (a Loop page tab, e.g. channel "Pages"), onenote, files, website, wiki or other. readable: true means teams_get_notes can return the tab's text; otherwise reason says where its content lives. A chat's Notes tab is a Loop page that Teams announces with a ThreadActivity/FileAdd event; this tool finds it in the newest ${SEARCH_DEPTH} messages (notesSearch reports how far it looked). If Microsoft Graph cannot list app tabs, tabsError says why and any Notes tab is still returned. Read-only.`,
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'Chat or channel conversation ID (e.g. "19:abc@thread.v2" for a group chat, "19:abc@thread.tacv2" for a channel).',
      },
    },
    required: ['conversationId'],
  },
};

const getNotesToolDefinition: Tool = {
  name: 'teams_get_notes',
  description: `Read the text of a chat's Notes and of Loop page tabs (e.g. channel "Pages") as plain text: headings as #, list items as - or 1., checkboxes as [ ]/[x], table cells separated by |. Pass conversationId to read every readable tab of that chat or channel (narrow to one with tabId from teams_get_tabs), or url to read one Loop file (a tab's contentUrl, or a .loop link shared in a message). Tabs that are not Loop pages (OneNote, Files, websites, apps) are listed under notRead with their link and the reason. Each page reports its last editor and edit time; text longer than 100,000 characters is cut (truncated: true). Read-only: never edits the notes. Uses Microsoft Graph via the existing Teams session; if Graph is unavailable only this tool and teams_get_tabs are affected.`,
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'Chat or channel conversation ID whose Notes/Loop tabs to read.' },
      tabId: { type: 'string', description: 'Optional, with conversationId: read only this tab (an id from teams_get_tabs).' },
      url: { type: 'string', description: 'Instead of conversationId: a Loop file URL (SharePoint) to read.' },
    },
  },
};

async function handleGetTabs(input: z.infer<typeof GetTabsInputSchema>): Promise<ToolResult> {
  const result = await listConversationTabs(input.conversationId);
  return handleApiResult(result, value => ({
    ...value,
    count: value.tabs.length,
  }));
}

function notRead(tab: ConversationTab, reason?: string) {
  return { tabId: tab.id, name: tab.name, kind: tab.kind, link: tab.webUrl ?? tab.websiteUrl ?? tab.contentUrl, reason: reason ?? tab.reason };
}

async function handleGetNotes(input: z.infer<typeof GetNotesInputSchema>): Promise<ToolResult> {
  if (input.url) {
    const page = await readLoopPage(input.url);
    return handleApiResult(page, value => ({ pages: [value], notRead: [] }));
  }

  const listed = await listConversationTabs(input.conversationId as string);
  if (!listed.ok) return { success: false, error: listed.error };
  const { tabs, notesSearch, tabsError } = listed.value;

  const selected = input.tabId ? tabs.filter(tab => tab.id === input.tabId) : tabs;
  if (input.tabId && selected.length === 0) {
    return {
      success: false,
      error: createError(ErrorCode.NOT_FOUND, `No tab with id "${input.tabId}" in this conversation`, {
        retryable: false, suggestions: ['Call teams_get_tabs to list the tab IDs'],
      }),
    };
  }

  const pages = [];
  const skipped = [];
  for (const tab of selected) {
    if (!tab.readable || !tab.contentUrl) {
      skipped.push(notRead(tab));
      continue;
    }
    const page = await readLoopPage(tab.contentUrl);
    if (page.ok) pages.push({ tabId: tab.id, tabName: tab.name, kind: tab.kind, ...page.value });
    else skipped.push(notRead(tab, `${page.error.code}: ${page.error.message}`));
  }

  return {
    success: true,
    data: {
      conversationId: listed.value.conversationId,
      pages,
      notRead: skipped,
      ...(notesSearch ? { notesSearch } : {}),
      ...(tabsError ? { tabsError } : {}),
      ...(pages.length === 0 && skipped.length === 0
        ? { note: notesSearch && !notesSearch.reachedStart
          ? `No Notes or Loop tab found in the newest ${notesSearch.messagesScanned} messages; an older Notes page may exist.`
          : 'This conversation has no Notes or Loop tabs.' }
        : {}),
    },
  };
}

export const getTabsTool: RegisteredTool<typeof GetTabsInputSchema> = {
  definition: getTabsToolDefinition,
  schema: GetTabsInputSchema,
  handler: handleGetTabs,
};

export const getNotesTool: RegisteredTool<typeof GetNotesInputSchema> = {
  definition: getNotesToolDefinition,
  schema: GetNotesInputSchema,
  handler: handleGetNotes,
};

export const tabTools = [getTabsTool, getNotesTool];

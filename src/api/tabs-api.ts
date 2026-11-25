/**
 * Conversation tabs and the text of chat Notes / Loop pages.
 *
 * - Channel and chat tabs come from Microsoft Graph (`/teams/{groupId}/channels/{id}/tabs`,
 *   `/chats/{id}/tabs`; `TeamsTab.ReadWrite.All` on the Teams client's Graph token).
 * - A chat's Notes tab is not a Graph tab: Teams stores it as a Loop file
 *   (`loop_loopThread.loop`) in a SharePoint Embedded container and announces it with a
 *   chatsvc `ThreadActivity/FileAdd` event (type `loop`, subtype `loopthread`). We find
 *   that event in the chat history and read the file through Graph.
 * - Loop files (chat Notes, channel Pages, Loop tabs) are read as text through Graph's
 *   `content?format=html` conversion, then flattened by `loopHtmlToText()`.
 *
 * Read-only. Like file downloads, Graph is optional: no failure here surfaces as
 * AUTH_REQUIRED/AUTH_EXPIRED, so it can never trigger the server's auto-login.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { requireGraphTokenAsync, requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { getSkypeAuthHeaders, CHATSVC_API } from '../utils/api-config.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { buildThreadMessagesUrl, cursorFromBackwardLink, getConversationProperties } from './chatsvc-messaging.js';
import { toDriveItemUrl } from './graph-files-api.js';
import { parseFileAddEvent, isChatNotesTarget, loopHtmlToText, type FileAddTarget } from '../utils/parsers-loop.js';
import type { RawChatsvcMessage } from '../types/api-responses.js';
import { MAX_NOTES_TEXT_CHARS, NOTES_SCAN_MAX_PAGES, NOTES_SCAN_PAGE_SIZE } from '../constants.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';

/** What a tab holds, and so whether `teams_get_notes` can read it. */
export type TabKind = 'notes' | 'loop' | 'onenote' | 'files' | 'website' | 'wiki' | 'other';

export interface ConversationTab {
  id: string;
  name: string;
  kind: TabKind;
  app?: string;
  appId?: string;
  /** URL of the tab's content (for Loop/Notes: the SharePoint file URL). */
  contentUrl?: string;
  websiteUrl?: string;
  /** Teams deep link to the tab. */
  webUrl?: string;
  /** True when `teams_get_notes` can return this tab's text. */
  readable: boolean;
  /** Why the text cannot be read here, when `readable` is false. */
  reason?: string;
  createdAt?: string;
  createdBy?: string;
}

export interface ConversationTabs {
  conversationId: string;
  conversationType: 'channel' | 'chat';
  tabs: ConversationTab[];
  /** How far the chat history was searched for the Notes page (chats only). */
  notesSearch?: { messagesScanned: number; reachedStart: boolean; found: boolean };
  /** Set when Graph could not list the app tabs; the Notes tab may still be present. */
  tabsError?: string;
}

export interface LoopPage {
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: string;
  text: string;
  lineCount: number;
  truncated: boolean;
}

// Well-known Teams app IDs for built-in tab types.
const ONENOTE_APP = '0d820ecd-def2-4297-adad-78056cde7c78';
const CHANNEL_PAGES_APP = '5ae80e49-7ada-461a-a6bd-c5df2e0cdb06';
const FILES_APP = '3ed5b337-c2c9-4d5d-b7b4-84ff09a8fc1c';
const WEBSITE_APP = 'com.microsoft.teamspace.tab.web';
const WIKI_APP = 'com.microsoft.teamspace.tab.wiki';

const REASONS: Record<Exclude<TabKind, 'notes' | 'loop'>, string> = {
  onenote: 'OneNote notebook tab: its pages are not read by this tool. Open webUrl in Teams.',
  files: 'Files tab: list its files with teams_get_shared_files and save one with teams_download_file.',
  website: 'Website tab: external page, not Teams content. Open websiteUrl.',
  wiki: 'Wiki tab (retired by Microsoft): content is not readable through this tool.',
  other: 'App tab: its content lives in the app. Open webUrl in Teams.',
};

/** True for a SharePoint URL of a Loop/Fluid file, which Graph can render as HTML. */
export function isLoopUrl(url: string | undefined): boolean {
  if (!url || !toDriveItemUrl(url)) return false;
  try {
    const { pathname } = new URL(url);
    return /\.(loop|fluid)$/i.test(decodeURIComponent(pathname)) || pathname.startsWith('/:fl:/');
  } catch {
    return false;
  }
}

interface RawGraphTab {
  id?: string;
  displayName?: string;
  webUrl?: string;
  configuration?: { entityId?: string | null; contentUrl?: string | null; websiteUrl?: string | null };
  teamsApp?: { id?: string; displayName?: string };
}

export function classifyTab(raw: RawGraphTab): ConversationTab {
  const appId = raw.teamsApp?.id;
  const contentUrl = raw.configuration?.contentUrl ?? undefined;
  let kind: TabKind;
  if (isLoopUrl(contentUrl)) kind = 'loop';
  else if (appId === ONENOTE_APP) kind = 'onenote';
  else if (appId === FILES_APP) kind = 'files';
  else if (appId === WEBSITE_APP) kind = 'website';
  else if (appId === WIKI_APP) kind = 'wiki';
  else kind = 'other';
  // A Channel Pages tab whose content URL is missing cannot be resolved to a file.
  const readable = kind === 'loop';
  return {
    id: raw.id ?? '',
    name: raw.displayName ?? '',
    kind,
    app: raw.teamsApp?.displayName,
    appId,
    contentUrl,
    websiteUrl: raw.configuration?.websiteUrl ?? undefined,
    webUrl: raw.webUrl,
    readable,
    ...(readable ? {} : {
      reason: appId === CHANNEL_PAGES_APP ? 'Page tab without a Loop file URL.' : REASONS[kind as keyof typeof REASONS],
    }),
  };
}

export function notesTab(target: FileAddTarget): ConversationTab {
  return {
    id: `notes:${target.id}`,
    name: 'Notes',
    kind: 'notes',
    app: 'Loop',
    contentUrl: target.uri,
    readable: isLoopUrl(target.uri),
    ...(isLoopUrl(target.uri) ? {} : { reason: 'The Notes file is not at a SharePoint URL this tool can resolve.' }),
    createdAt: target.createdAt,
    createdBy: target.creatorMri,
  };
}

/**
 * GET from Microsoft Graph with the Teams client's token. Maps auth failures away
 * from AUTH_* codes (see the module comment) and discards a rejected token.
 */
async function graphGet<T>(url: string, accept = 'application/json'): Promise<Result<T>> {
  const token = await requireGraphTokenAsync();
  if (!token.ok) return token;
  const response = await httpRequest<T>(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token.value}`, Accept: accept },
  });
  if (response.ok) return ok(response.value.data);
  const { code, message } = response.error;
  if (code === ErrorCode.AUTH_EXPIRED) {
    invalidateAccessToken(token.value);
    return err(createError(ErrorCode.API_ERROR,
      'Microsoft Graph rejected its access token; it has been discarded, so a retry will request a new one.',
      { retryable: true }));
  }
  if (code === ErrorCode.AUTH_REQUIRED) {
    return err(createError(ErrorCode.ACCESS_DENIED, `Microsoft Graph refused access: ${message}`, { retryable: false }));
  }
  return response;
}

/** Strips a channel reply suffix (`;messageid=…`) so the ID names the conversation itself. */
function baseConversationId(conversationId: string): string {
  return conversationId.split(';')[0];
}

/**
 * Searches the chat history, newest first, for `ThreadActivity/FileAdd` events of
 * Loop files. The chat's Notes page is the newest `loopthread` one.
 */
async function findChatNotes(conversationId: string): Promise<Result<{
  target?: FileAddTarget; messagesScanned: number; reachedStart: boolean;
}>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;
  const endpoint = CHATSVC_API.messages(region, conversationId, undefined, baseUrl);

  let cursor: string | undefined;
  let messagesScanned = 0;
  for (let page = 0; page < NOTES_SCAN_MAX_PAGES; page++) {
    const response = await httpRequest<{ messages?: RawChatsvcMessage[]; _metadata?: { backwardLink?: string } }>(
      buildThreadMessagesUrl(endpoint, { limit: NOTES_SCAN_PAGE_SIZE, cursor }),
      { method: 'GET', headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl) },
    );
    if (!response.ok) return response;
    const messages = response.value.data.messages ?? [];
    messagesScanned += messages.length;
    for (const message of messages) {
      if (message.messagetype !== 'ThreadActivity/FileAdd' || !message.content) continue;
      const target = parseFileAddEvent(message.content);
      if (target && isChatNotesTarget(target)) return ok({ target, messagesScanned, reachedStart: false });
    }
    cursor = cursorFromBackwardLink(response.value.data._metadata?.backwardLink);
    if (!cursor || messages.length === 0) return ok({ messagesScanned, reachedStart: true });
  }
  return ok({ messagesScanned, reachedStart: false });
}

/**
 * Lists a conversation's tabs: Graph app tabs for channels and chats, plus a chat's
 * Notes page (found in its history). A Graph failure is reported in `tabsError`
 * rather than hiding a Notes tab that was found.
 */
export async function listConversationTabs(conversationId: string): Promise<Result<ConversationTabs>> {
  const id = baseConversationId(conversationId);
  const props = await getConversationProperties(id);
  if (!props.ok) return props;
  const groupId = props.value.groupId;
  const isChannel = props.value.conversationType === 'Channel' && !!groupId;

  const tabsUrl = isChannel
    ? `${GRAPH}/teams/${encodeURIComponent(groupId)}/channels/${encodeURIComponent(id)}/tabs?$expand=teamsApp`
    : `${GRAPH}/chats/${encodeURIComponent(id)}/tabs?$expand=teamsApp`;
  const graphTabs = await graphGet<{ value?: RawGraphTab[] }>(tabsUrl);

  const tabs: ConversationTab[] = [];
  let notesSearch: ConversationTabs['notesSearch'];
  if (!isChannel) {
    const notes = await findChatNotes(id);
    if (!notes.ok) return notes;
    notesSearch = {
      messagesScanned: notes.value.messagesScanned,
      reachedStart: notes.value.reachedStart,
      found: !!notes.value.target,
    };
    if (notes.value.target) tabs.push(notesTab(notes.value.target));
  }
  if (graphTabs.ok) tabs.push(...(graphTabs.value.value ?? []).map(classifyTab));

  return ok({
    conversationId: id,
    conversationType: isChannel ? 'channel' : 'chat',
    tabs,
    ...(notesSearch ? { notesSearch } : {}),
    ...(graphTabs.ok ? {} : { tabsError: describe(graphTabs.error) }),
  });
}

function describe(error: McpError): string {
  return `${error.code}: ${error.message}`.slice(0, 300);
}

interface RawLoopItem {
  id?: string;
  name?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  file?: unknown;
  parentReference?: { driveId?: string };
}

/**
 * Reads a Loop/Fluid file (chat Notes, channel Page, Loop tab) as plain text.
 * `url` is the SharePoint file URL (a tab's contentUrl) or a Graph drive-item URL.
 */
export async function readLoopPage(url: string): Promise<Result<LoopPage>> {
  const itemUrl = toDriveItemUrl(url);
  if (!itemUrl) {
    return err(createError(ErrorCode.INVALID_INPUT,
      'url must be a SharePoint Loop file URL (a Notes or Loop tab contentUrl from teams_get_tabs) or a Graph drive-item URL'));
  }
  const select = new URLSearchParams({ '$select': 'id,name,webUrl,lastModifiedDateTime,lastModifiedBy,file,parentReference' });
  const meta = await graphGet<RawLoopItem>(`${itemUrl}?${select}`);
  if (!meta.ok) {
    if (meta.error.code === ErrorCode.NOT_FOUND) {
      return err(createError(ErrorCode.NOT_FOUND, 'Loop page not found: it may have been deleted, or you no longer have access.', { retryable: false }));
    }
    return meta;
  }
  const item = meta.value;
  const driveId = item.parentReference?.driveId;
  if (!item.file || !item.id || !driveId) {
    return err(createError(ErrorCode.INVALID_INPUT, 'This URL is not a file', { retryable: false }));
  }
  if (!/\.(loop|fluid)$/i.test(item.name ?? '')) {
    return err(createError(ErrorCode.INVALID_INPUT,
      `"${item.name}" is not a Loop page; save it with teams_download_file instead.`, { retryable: false }));
  }

  const contentUrl = `${GRAPH}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}/content?format=html`;
  const html = await graphGet<string>(contentUrl, 'text/html');
  if (!html.ok) return html;
  if (typeof html.value !== 'string') {
    return err(createError(ErrorCode.API_ERROR, 'Microsoft Graph did not return an HTML rendering of this Loop page', { retryable: false }));
  }

  const fullText = loopHtmlToText(html.value);
  const truncated = fullText.length > MAX_NOTES_TEXT_CHARS;
  const text = truncated ? fullText.slice(0, MAX_NOTES_TEXT_CHARS) : fullText;
  return ok({
    name: item.name ?? '',
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    lastModifiedBy: item.lastModifiedBy?.user?.displayName,
    text,
    lineCount: text ? text.split('\n').filter(line => line.trim()).length : 0,
    truncated,
  });
}

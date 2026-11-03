/**
 * File-related tool handlers: list files shared in a conversation, and
 * download a file (shared file or assignment attachment) via Microsoft Graph.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { getSharedFiles } from '../api/files-api.js';
import { downloadFile } from '../api/graph-files-api.js';
import { saveMessageAttachments } from '../api/attachments-api.js';
import {
  DEFAULT_FILES_PAGE_SIZE,
  MAX_FILES_PAGE_SIZE,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetSharedFilesInputSchema = z.object({
  conversationId: z.string().min(1),
  pageSize: z.number().min(1).max(MAX_FILES_PAGE_SIZE).optional().default(DEFAULT_FILES_PAGE_SIZE),
  skipToken: z.string().optional(),
});

export const DownloadAttachmentInputSchema = z.object({
  conversationId: z.string().min(1),
  messageId: z.string().min(1),
  attachmentId: z.string().min(1).optional(),
});

export const DownloadFileInputSchema = z.object({
  url: z.string().url(),
  outputPath: z.string().min(1),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getSharedFilesToolDefinition: Tool = {
  name: 'teams_get_shared_files',
  description: 'Get files and links shared in a Teams conversation. Returns file names, URLs, extensions, sizes, and who shared them. Works for channels, group chats, 1:1 chats, and meeting chats. Use the conversationId from other tools (teams_get_favorites, teams_search, teams_find_channel, teams_get_chat). Supports pagination via skipToken for conversations with many files. To save the contents of a File item locally, pass its webUrl to teams_download_file.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get shared files for (e.g., "19:abc@thread.tacv2" for a channel, or a chat conversation ID).',
      },
      pageSize: {
        type: 'number',
        description: `Number of items per page (default: ${DEFAULT_FILES_PAGE_SIZE}, max: ${MAX_FILES_PAGE_SIZE})`,
      },
      skipToken: {
        type: 'string',
        description: 'Continuation token from a previous response to get the next page of results.',
      },
    },
    required: ['conversationId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetSharedFiles(
  input: z.infer<typeof GetSharedFilesInputSchema>
): Promise<ToolResult> {
  const result = await getSharedFiles(input.conversationId, {
    pageSize: input.pageSize,
    skipToken: input.skipToken,
  });

  return handleApiResult(result, (value) => ({
    conversationId: value.conversationId,
    returned: value.returned,
    files: value.files,
    ...(value.skipToken ? { skipToken: value.skipToken, hasMore: true } : { hasMore: false }),
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getSharedFilesTool: RegisteredTool<typeof GetSharedFilesInputSchema> = {
  definition: getSharedFilesToolDefinition,
  schema: GetSharedFilesInputSchema,
  handler: handleGetSharedFiles,
};

const downloadFileToolDefinition: Tool = {
  name: 'teams_download_file',
  description:
    "Download a Teams file to a local path. Accepts the webUrl of a File item from teams_get_shared_files (SharePoint/OneDrive: chat uploads and channel files, including Doc.aspx viewer links), or the fileUrl of an assignment attachment or submission file from teams_get_assignment (a Microsoft Graph drive-item URL). Link items, Microsoft Forms and non-SharePoint URLs are not files and cannot be downloaded. Saves the raw file to an absolute path on the MCP server machine and returns its name, size, content type and SHA-256. The parent directory must exist and existing files are never overwritten. Streams to disk with no size limit; a transfer that stalls for 30 seconds is cancelled and the partial file removed. Respects the owner's download block. Uses Microsoft Graph via the existing Teams session (no extra sign-in); if Graph is unavailable only this tool fails, and no other Teams tool is affected.",
  inputSchema: {
    type: 'object',
    properties: {
      url: { type: 'string', description: "A File item's webUrl from teams_get_shared_files, or an attachment's fileUrl from teams_get_assignment." },
      outputPath: { type: 'string', description: 'Absolute destination file path on the MCP server machine. Must not already exist; the parent directory must exist.' },
    },
    required: ['url', 'outputPath'],
  },
};

async function handleDownloadFile(
  input: z.infer<typeof DownloadFileInputSchema>
): Promise<ToolResult> {
  const result = await downloadFile(input.url, input.outputPath);
  return handleApiResult(result, (value) => ({ ...value }));
}

export const downloadFileTool: RegisteredTool<typeof DownloadFileInputSchema> = {
  definition: downloadFileToolDefinition,
  schema: DownloadFileInputSchema,
  handler: handleDownloadFile,
};

const downloadAttachmentToolDefinition: Tool = {
  name: 'teams_download_attachment',
  description:
    "Save the images and files attached to a Teams message to a private local folder and return their paths, so you can open them (for example with a file Read tool to view a screenshot). Use the attachments listed by teams_get_thread or teams_get_message; an image-only message has empty content but lists its image there. Covers images pasted into chats and files uploaded to chats or channels. Read-only: nothing is sent to Teams. Files go to the Teams MCP data directory (owner-only permissions); a repeated request reuses the earlier file (cached: true). If one attachment fails, the others are still saved and the failure is reported next to them.",
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation ID containing the message.' },
      messageId: { type: 'string', description: 'The numeric message ID (id from teams_get_thread or teams_get_message).' },
      attachmentId: { type: 'string', description: 'Optional: only save this attachment (an attachments[].id). Omit to save all of the message\'s attachments.' },
    },
    required: ['conversationId', 'messageId'],
  },
};

async function handleDownloadAttachment(
  input: z.infer<typeof DownloadAttachmentInputSchema>
): Promise<ToolResult> {
  const result = await saveMessageAttachments(input.conversationId, input.messageId, input.attachmentId);
  return handleApiResult(result, (value) => ({ ...value }));
}

export const downloadAttachmentTool: RegisteredTool<typeof DownloadAttachmentInputSchema> = {
  definition: downloadAttachmentToolDefinition,
  schema: DownloadAttachmentInputSchema,
  handler: handleDownloadAttachment,
};

/** All file-related tools. */
export const fileTools = [getSharedFilesTool, downloadFileTool, downloadAttachmentTool];

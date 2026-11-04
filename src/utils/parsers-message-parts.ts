import type { RawChatsvcMessage } from '../types/api-responses.js';
import { stripHtml } from './parsers-html.js';

export interface MessageAttachment {
  id: string;
  kind: 'image' | 'file';
  source: 'teams-media' | 'sharepoint';
  name?: string;
  type?: string;
  width?: number;
  height?: number;
  url: string;
}

// Only these hosts ever receive the Skype token.
const TEAMS_MEDIA_URL = /^https:\/\/[a-z0-9-]+\.asm\.skype\.com\/v1\/objects\/([A-Za-z0-9_-]+)(?:\/views\/[A-Za-z0-9_]+)?$/;

export function isTeamsMediaUrl(url: string): boolean {
  return TEAMS_MEDIA_URL.test(url);
}

function attributes(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const [, name, value] of tag.matchAll(/([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*"([^"]*)"/g)) {
    attrs[name.toLowerCase()] = value.replace(/&amp;/g, '&');
  }
  return attrs;
}

function dimension(value: string | undefined): number | undefined {
  const parsed = value ? Math.round(Number(value)) : NaN;
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function inlineImages(content: string): MessageAttachment[] {
  const images: MessageAttachment[] = [];
  for (const [tag] of content.matchAll(/<img\b[^>]*>/gi)) {
    const attrs = attributes(tag);
    const objectId = attrs.src?.match(TEAMS_MEDIA_URL)?.[1];
    if (!objectId) continue;
    images.push({
      id: objectId,
      kind: 'image',
      source: 'teams-media',
      type: attrs.itemscope || undefined,
      width: dimension(attrs.width),
      height: dimension(attrs.height),
      url: attrs.src,
    });
  }
  return images;
}

function uriObjects(content: string): MessageAttachment[] {
  const objects: MessageAttachment[] = [];
  for (const [whole, open] of content.matchAll(/(<URIObject\b[^>]*>)[\s\S]*?<\/URIObject>/gi)) {
    const attrs = attributes(open);
    const objectId = attrs.uri?.match(TEAMS_MEDIA_URL)?.[1];
    if (!objectId) continue;
    const name = whole.match(/<OriginalName\b[^>]*\bv="([^"]*)"/i)?.[1];
    objects.push({
      id: objectId,
      kind: /^Picture/i.test(attrs.type ?? '') ? 'image' : 'file',
      source: 'teams-media',
      name: name ? stripHtml(name) : undefined,
      url: attrs.uri,
    });
  }
  return objects;
}

const text = (value: unknown): string | undefined => (typeof value === 'string' && value ? value : undefined);

function sharedFiles(files: unknown): MessageAttachment[] {
  let list: unknown = files;
  if (typeof files === 'string') {
    try { list = JSON.parse(files); } catch { return []; }
  }
  if (!Array.isArray(list)) return [];
  return list.flatMap((raw: unknown, index): MessageAttachment[] => {
    if (typeof raw !== 'object' || raw === null) return [];
    const file = raw as Record<string, unknown>;
    const fileInfo = typeof file.fileInfo === 'object' && file.fileInfo !== null ? file.fileInfo as Record<string, unknown> : {};
    const url = text(file.objectUrl) ?? text(fileInfo.fileUrl);
    if (!url?.startsWith('https://')) return [];
    return [{
      id: text(file.itemid) ?? text(file.id) ?? `file-${index + 1}`,
      kind: 'file',
      source: 'sharepoint',
      name: text(file.fileName) ?? text(file.title),
      type: text(file.fileType),
      url,
    }];
  });
}

export function extractAttachments(msg: RawChatsvcMessage): MessageAttachment[] {
  const content = msg.content ?? '';
  const found = [...inlineImages(content), ...uriObjects(content), ...sharedFiles(msg.properties?.files)];
  const seen = new Set<string>();
  return found.filter(attachment => !seen.has(attachment.id) && seen.add(attachment.id));
}

export interface SystemEvent {
  type: string;
  summary: string;
}

export function isSystemMessageType(messageType: string): boolean {
  return messageType.startsWith('Event/')
    || messageType.startsWith('ThreadActivity/')
    || messageType.startsWith('RichText/Media_Call');
}

function plural(count: number, noun: string): string {
  return `${count} ${noun}${count === 1 ? '' : 's'}`;
}

function formatDuration(seconds: number): string {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return [h && `${h}h`, m && `${m}m`, (s || (!h && !m)) && `${s}s`].filter(Boolean).join(' ');
}

function callEvent(content: string): SystemEvent {
  const names = [...content.matchAll(/<displayName>([\s\S]*?)<\/displayName>/gi)]
    .map(([, name]) => stripHtml(name)).filter(Boolean);
  const durations = [...content.matchAll(/<duration>(\d+)<\/duration>/gi)].map(([, d]) => Number(d));
  const longest = durations.length ? Math.max(...durations) : 0;
  const people = names.length ? ` with ${plural(names.length, 'participant')}: ${names.join(', ')}` : '';
  const eventType = content.match(/<callEventType>\s*(\w+)\s*<\/callEventType>/i)?.[1]?.toLowerCase();
  if (/<ended\s*\/>/i.test(content) || eventType === 'callended') {
    return { type: 'call_ended', summary: `Call ended${longest ? ` after ${formatDuration(longest)}` : ''}${people}` };
  }
  if (/<started\s*\/>/i.test(content) || eventType === 'callstarted') return { type: 'call_started', summary: 'Call started' };
  if (/<missed\s*\/>/i.test(content)) return { type: 'call_missed', summary: 'Missed call' };
  return { type: 'call', summary: `Call event${people}` };
}

function membershipEvent(messageType: string, content: string): SystemEvent {
  const verb: Record<string, [string, string]> = {
    'ThreadActivity/MemberJoined': ['members_joined', 'joined'],
    'ThreadActivity/MemberLeft': ['members_left', 'left'],
    'ThreadActivity/AddMember': ['members_added', 'added'],
    'ThreadActivity/DeleteMember': ['members_removed', 'removed'],
  };
  const [type, action] = verb[messageType];
  let names: string[] = [];
  let count = 0;
  try {
    const members = (JSON.parse(content) as { members?: Array<{ friendlyname?: string }> }).members ?? [];
    count = members.length;
    names = members.map(m => m.friendlyname?.trim() ?? '').filter(Boolean);
  } catch {
    count = (content.match(/<target>/gi) ?? []).length;
  }
  if (count === 0) return { type, summary: `Members ${action}` };
  const who = names.length === count ? `: ${names.join(', ')}` : '';
  return { type, summary: `${plural(count, 'member')} ${action}${who}` };
}

export function parseSystemEvent(messageType: string, content: string): SystemEvent | null {
  if (!isSystemMessageType(messageType)) return null;
  if (messageType === 'Event/Call') return callEvent(content);
  if (messageType === 'ThreadActivity/TopicUpdate') {
    const topic = content.match(/<value>([\s\S]*?)<\/value>/i)?.[1];
    return { type: 'topic_changed', summary: topic ? `Chat renamed to "${stripHtml(topic)}"` : 'Chat renamed' };
  }
  if (/^ThreadActivity\/(MemberJoined|MemberLeft|AddMember|DeleteMember)$/.test(messageType)) {
    return membershipEvent(messageType, content);
  }
  if (messageType === 'RichText/Media_CallRecording') return { type: 'call_recording', summary: 'Call recording available' };
  if (messageType === 'RichText/Media_CallTranscript') return { type: 'call_transcript', summary: 'Call transcript available' };
  const name = messageType.split('/').pop() ?? messageType;
  return { type: name.replace(/([a-z])([A-Z])/g, '$1_$2').toLowerCase(), summary: `Teams event: ${messageType}` };
}

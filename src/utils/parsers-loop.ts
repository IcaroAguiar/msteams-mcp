/** Parsers for the `ThreadActivity/FileAdd` event and Graph's Loop HTML rendering. */

import { HTML_ENTITIES } from './html-entities.js';

/** A file attached to a conversation, as recorded by a `ThreadActivity/FileAdd` event. */
export interface FileAddTarget {
  id: string;
  uri: string;
  type: string;
  subtype?: string;
  creatorMri?: string;
  createdAt?: string;
}

function tagText(xml: string, tag: string): string | undefined {
  const match = xml.match(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`));
  return match ? decodeEntities(match[1]).trim() : undefined;
}

/** Parses the `<fileadd>` XML of a `ThreadActivity/FileAdd` event; null when it has no usable target. */
export function parseFileAddEvent(content: string): FileAddTarget | null {
  const target = content.match(/<target>([\s\S]*?)<\/target>/)?.[1];
  if (!target) return null;
  const id = tagText(target, 'id');
  const uri = tagText(target, 'uri');
  const type = tagText(target, 'type');
  if (!id || !uri || !type) return null;
  const createdAtMs = Number(tagText(target, 'createdAt'));
  return {
    id,
    uri,
    type,
    subtype: tagText(target, 'subtype') || undefined,
    creatorMri: tagText(target, 'creator') || undefined,
    createdAt: Number.isFinite(createdAtMs) && createdAtMs > 0 ? new Date(createdAtMs).toISOString() : undefined,
  };
}

/** True for the Loop page Teams creates as a chat's Notes tab. */
export function isChatNotesTarget(target: FileAddTarget): boolean {
  return target.type === 'loop' && target.subtype?.toLowerCase() === 'loopthread';
}

// HTML maps numeric references in 0x80–0x9F to their windows-1252 characters.
const C1_REPLACEMENTS: Record<number, number> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021,
  0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018,
  0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc,
  0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178,
};

function decodeCodePoint(code: number): string {
  if (!Number.isFinite(code) || code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return '\uFFFD';
  return String.fromCodePoint(C1_REPLACEMENTS[code] ?? code);
}

/**
 * Decodes HTML character references. A legacy name without its semicolon is decoded only
 * when it is not followed by `=`, so query strings such as `?a=1&copy=2` survive.
 */
export function decodeEntities(text: string): string {
  return text.replace(/&(?:#[xX]([0-9a-fA-F]+);?|#(\d+);?|([a-zA-Z][a-zA-Z0-9]*)(;?))(=?)/g,
    (whole, hex: string | undefined, dec: string | undefined, name: string | undefined, semicolon: string, equals: string) => {
      if (hex !== undefined) return decodeCodePoint(parseInt(hex, 16)) + equals;
      if (dec !== undefined) return decodeCodePoint(parseInt(dec, 10)) + equals;
      if (semicolon) return (HTML_ENTITIES[`${name};`] ?? `&${name};`) + equals;
      return !equals && Object.hasOwn(HTML_ENTITIES, name as string) ? HTML_ENTITIES[name as string] : whole;
    });
}

const BLOCK_TAGS = new Set(['p', 'section', 'article', 'blockquote', 'pre', 'table', 'tr', 'ul', 'ol', 'hr']);

/** Loop HTML to text: `#` headings, `-`/`1.` bullets by depth, `[ ]`/`[x]`, `|` cells. */
export function loopHtmlToText(html: string, { cut = false }: { cut?: boolean } = {}): string {
  let source = html
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(script|style|head)\b[^>]*>[\s\S]*?<\/\1>/gi, '');
  // Only a cut body can end inside a comment, a script or style block, or a tag; on a complete page
  // the same pattern would also match `<script` inside an attribute value.
  if (cut) source = source.replace(/<!--[\s\S]*$|<(?:script|style)\b[\s\S]*$|<[^>]*$/i, '');

  const lists: Array<{ ordered: boolean; next: number }> = [];
  // Loop wraps @mentions in `<div style="display: inline">`, which must not break the line.
  const divIsInline: boolean[] = [];
  let out = '';
  // Inside a table cell a block break becomes a space, so each row stays on one line.
  let cellDepth = 0;
  const newline = () => {
    if (cellDepth > 0) {
      if (out && !/[ \n]$/.test(out)) out += ' ';
    } else if (out && !out.endsWith('\n')) {
      out += '\n';
    }
  };

  for (const token of source.matchAll(/<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>|([^<]+)/g)) {
    const [, closing, rawName, attrs, text] = token;
    if (text !== undefined) {
      const words = decodeEntities(text).replace(/[ \t\r\n\f]+/g, ' ');
      out += !out || out.endsWith('\n') ? words.trimStart() : words;
      continue;
    }
    const name = rawName.toLowerCase();
    const heading = /^h([1-6])$/.exec(name);
    if (heading) {
      newline();
      if (!closing) out += `${'#'.repeat(Number(heading[1]))} `;
      continue;
    }
    switch (name) {
      case 'div': {
        const inline = closing ? divIsInline.pop() : /display\s*:\s*inline\b/i.test(attrs);
        if (!closing) divIsInline.push(inline ?? false);
        if (!inline) newline();
        break;
      }
      case 'br':
        if (cellDepth > 0) newline();
        else out += '\n';
        break;
      case 'ul':
      case 'ol':
        newline();
        if (closing) lists.pop();
        else lists.push({ ordered: name === 'ol', next: 1 });
        break;
      case 'li':
        newline();
        if (!closing) {
          const list = lists[lists.length - 1];
          const indent = '  '.repeat(Math.max(lists.length - 1, 0));
          out += list?.ordered ? `${indent}${list.next++}. ` : `${indent}- `;
        }
        break;
      case 'td':
      case 'th':
        if (closing) {
          cellDepth = Math.max(cellDepth - 1, 0);
          out += ' | ';
        } else {
          cellDepth++;
        }
        break;
      case 'input':
        if (/type\s*=\s*["']?checkbox/i.test(attrs)) out += /\bchecked\b/i.test(attrs) ? '[x] ' : '[ ] ';
        break;
      default:
        if (BLOCK_TAGS.has(name)) newline();
    }
  }

  return out
    .split('\n')
    .map(line => line.replace(/(\S)[ \t]+/g, '$1 ').replace(/\s*\|\s*$/, '').trimEnd())
    .map(line => (/^\s*(?:[-#]+|\d+\.)\s*$/.test(line) ? '' : line))
    .join('\n')
    .replace(/^\n+|\n+$/g, '')
    .replace(/\n{3,}/g, '\n\n');
}

/** Parsers for the `ThreadActivity/FileAdd` event and Graph's Loop HTML rendering. */

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

const NAMED_ENTITIES: Record<string, string> = {
  nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'",
};

export function decodeEntities(text: string): string {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, entity: string) => {
    if (entity[0] === '#') {
      const code = entity[1] === 'x' || entity[1] === 'X' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    }
    return NAMED_ENTITIES[entity.toLowerCase()] ?? whole;
  });
}

const BLOCK_TAGS = new Set(['p', 'section', 'article', 'blockquote', 'pre', 'table', 'tr', 'ul', 'ol', 'hr']);

/** Loop HTML to text: `#` headings, `-`/`1.` bullets by depth, `[ ]`/`[x]`, `|` cells. */
export function loopHtmlToText(html: string): string {
  const source = html
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(script|style|head)\b[^>]*>[\s\S]*?<\/\1>/gi, '');

  const lists: Array<{ ordered: boolean; next: number }> = [];
  // Loop wraps @mentions in `<div style="display: inline">`, which must not break the line.
  const divIsInline: boolean[] = [];
  let out = '';
  const newline = () => { if (out && !out.endsWith('\n')) out += '\n'; };

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
        out += '\n';
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
        if (closing) out += ' | ';
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

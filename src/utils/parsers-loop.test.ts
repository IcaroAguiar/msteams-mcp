import { describe, expect, it } from 'vitest';
import { isChatNotesTarget, loopHtmlToText, parseFileAddEvent } from './parsers-loop.js';

const NOTES_EVENT = '<fileadd><eventtime>1700000000001</eventtime><initiator>8:orgid:aaaa</initiator><target>'
  + '<id>11111111-2222-3333-4444-555555555555</id>'
  + '<uri>https://contoso.sharepoint.com/contentstorage/CSP_x/Document%20Library/19_chat_thread_v2/loop_loopThread.loop?nav=abc&amp;x=1</uri>'
  + '<type>loop</type><subtype>loopthread</subtype><creator>8:orgid:aaaa</creator><createdAt>1700000000000</createdAt>'
  + '</target></fileadd>';

describe('parseFileAddEvent', () => {
  it('reads the Notes page target, decoding the URI', () => {
    const target = parseFileAddEvent(NOTES_EVENT);
    expect(target).toEqual({
      id: '11111111-2222-3333-4444-555555555555',
      uri: 'https://contoso.sharepoint.com/contentstorage/CSP_x/Document%20Library/19_chat_thread_v2/loop_loopThread.loop?nav=abc&x=1',
      type: 'loop',
      subtype: 'loopthread',
      creatorMri: '8:orgid:aaaa',
      createdAt: new Date(1700000000000).toISOString(),
    });
    expect(isChatNotesTarget(target!)).toBe(true);
  });

  it('does not take other Loop files or ordinary file uploads for Notes', () => {
    const page = parseFileAddEvent(NOTES_EVENT.replace('<subtype>loopthread</subtype>', '<subtype>page</subtype>'));
    expect(page && isChatNotesTarget(page)).toBe(false);
    const file = parseFileAddEvent(NOTES_EVENT.replace('<type>loop</type>', '<type>docx</type>'));
    expect(file && isChatNotesTarget(file)).toBe(false);
  });

  it('returns null without a complete target', () => {
    expect(parseFileAddEvent('<fileadd><eventtime>1</eventtime></fileadd>')).toBeNull();
    expect(parseFileAddEvent('<fileadd><target><id>x</id><type>loop</type></target></fileadd>')).toBeNull();
  });
});

describe('loopHtmlToText', () => {
  it('keeps the structure of a Loop page: headings, bullets, mentions and paragraphs', () => {
    const html = '<html><meta name="GENERATOR" content="Microsoft Loop"><body><div style="margin: 25px"><div lang="en-gb">'
      + '<ul style="margin-top: 0px"><li><span>First point &amp; detail</span></li></ul>'
      + '<h2 id="a"><span>Section one</span></h2><h2 id="b"><span>Section two</span></h2>'
      + '<div data-docparser-context="paragraph"><span></span><div style="display: inline;"><a href="mailto:a@example.com">Alex Doe</a></div><span> please review</span></div>'
      + '<div data-docparser-context="paragraph"><span></span><br/></div>'
      + '</div></div></body></html>';
    expect(loopHtmlToText(html)).toBe([
      '- First point & detail',
      '## Section one',
      '## Section two',
      'Alex Doe please review',
    ].join('\n'));
  });

  it('numbers ordered lists, indents nested lists and renders checkboxes', () => {
    const html = '<ol><li>One<ul><li><input type="checkbox" checked>Done</li><li><input type="checkbox">Todo</li></ul></li><li>Two</li></ol>';
    expect(loopHtmlToText(html)).toBe(['1. One', '  - [x] Done', '  - [ ] Todo', '2. Two'].join('\n'));
  });

  it('keeps each table row on one line with its cells separated, and decodes every entity form', () => {
    const html = '<table><tr><th>Name</th><th>Status</th></tr>'
      + '<tr><td><p>Caf&eacute;</p><p>&copy 2026&hellip;</p></td><td><p>&#x2705; ok</p><p>a&#233;b &NotEqualTilde;</p></td></tr></table>'
      + '<p>?a=1&copy=2</p>';
    expect(loopHtmlToText(html)).toBe(['Name | Status', 'Café © 2026… | ✅ ok aéb \u2242\u0338', '?a=1&copy=2'].join('\n'));
  });

  it('returns an empty string for an empty page', () => {
    expect(loopHtmlToText('<html><meta name="GENERATOR" content="Microsoft Loop"><body><div><div lang="pt-br"></div></div></body></html>')).toBe('');
  });
});

import { describe, expect, it } from 'vitest';
import { extractAttachments, parseSystemEvent } from './parsers-message-parts.js';

const IMAGE_ID = '0-wus-d1-0123456789abcdef0123456789abcdef';
const IMAGE_URL = `https://us-api.asm.skype.com/v1/objects/${IMAGE_ID}/views/imgo`;

describe('extractAttachments', () => {
  it('lists an image pasted into an otherwise empty message', () => {
    const content = `<p><img src="${IMAGE_URL}" itemtype="http://schema.skype.com/AMSImage" itemscope="png" width="321.79" height="250" alt="image" itemid="${IMAGE_ID}"></p>`;
    expect(extractAttachments({ content })).toEqual([{
      id: IMAGE_ID, kind: 'image', source: 'teams-media', type: 'png', width: 322, height: 250, url: IMAGE_URL,
    }]);
  });

  it('lists files from the files property, whether serialized or not', () => {
    const file = {
      itemid: 'file-a', fileName: 'notes.md', fileType: 'md',
      objectUrl: 'https://contoso-my.sharepoint.com/personal/someone/Documents/notes.md',
      fileInfo: { fileUrl: 'https://contoso-my.sharepoint.com/ignored.md' },
    };
    const expected = [{ id: 'file-a', kind: 'file', source: 'sharepoint', name: 'notes.md', type: 'md', url: file.objectUrl }];
    expect(extractAttachments({ content: '', properties: { files: JSON.stringify([file]) } })).toEqual(expected);
    expect(extractAttachments({ content: '', properties: { files: [file] } })).toEqual(expected);
  });

  it('falls back to fileInfo.fileUrl and skips files without an HTTPS URL', () => {
    const files = JSON.stringify([
      { fileName: 'a.pdf', fileInfo: { fileUrl: 'https://contoso.sharepoint.com/a.pdf' } },
      { fileName: 'b.pdf', objectUrl: 'http://contoso.sharepoint.com/b.pdf' },
    ]);
    expect(extractAttachments({ properties: { files } })).toEqual([
      { id: 'file-1', kind: 'file', source: 'sharepoint', name: 'a.pdf', type: undefined, url: 'https://contoso.sharepoint.com/a.pdf' },
    ]);
  });

  it('ignores emoji, external images and media URLs on other hosts', () => {
    const content = [
      '<img itemtype="http://schema.skype.com/Emoji" src="https://statics.teams.cdn.office.net/evergreen-assets/emoji.png">',
      '<img src="https://example.com/cat.png">',
      `<img itemtype="http://schema.skype.com/AMSImage" src="https://asm.skype.com.evil.example/v1/objects/${IMAGE_ID}/views/imgo">`,
    ].join('');
    expect(extractAttachments({ content })).toEqual([]);
  });

  it('reads legacy URIObject media', () => {
    const content = `<URIObject type="Picture.1" uri="https://eu-api.asm.skype.com/v1/objects/${IMAGE_ID}"><OriginalName v="shot.png"/></URIObject>`;
    expect(extractAttachments({ content })).toEqual([{
      id: IMAGE_ID, kind: 'image', source: 'teams-media', name: 'shot.png',
      url: `https://eu-api.asm.skype.com/v1/objects/${IMAGE_ID}`,
    }]);
  });

  it('lists an image referenced twice only once', () => {
    const img = `<img src="${IMAGE_URL}" itemtype="http://schema.skype.com/AMSImage">`;
    expect(extractAttachments({ content: img + img })).toHaveLength(1);
  });

  it('tolerates a malformed files property', () => {
    expect(extractAttachments({ content: 'hello', properties: { files: '{not json' } })).toEqual([]);
  });
});

describe('parseSystemEvent', () => {
  it('returns null for messages written by people', () => {
    expect(parseSystemEvent('RichText/Html', '<p>hi</p>')).toBeNull();
    expect(parseSystemEvent('Text', 'hi')).toBeNull();
  });

  it('summarizes an ended call with its duration and participants', () => {
    const part = (name: string) => `<part identity="8:orgid:x"><name>8:orgid:x</name><displayName>${name}</displayName><duration>1957</duration></part>`;
    const content = `<ended/><partlist alt="" count="2">${part('Ada Lovelace')}${part('Alan Turing')}</partlist>`;
    expect(parseSystemEvent('Event/Call', content)).toEqual({
      type: 'call_ended', summary: 'Call ended after 32m 37s with 2 participants: Ada Lovelace, Alan Turing',
    });
  });

  it('summarizes a started call', () => {
    expect(parseSystemEvent('Event/Call', '<partlist type="started" alt=""><started/></partlist>')).toEqual({ type: 'call_started', summary: 'Call started' });
    expect(parseSystemEvent('Event/Call', '<partlist alt=""></partlist><meetingDetails></meetingDetails><callEventType>callStarted</callEventType>'))
      .toEqual({ type: 'call_started', summary: 'Call started' });
  });

  it('summarizes a topic change, decoding entities', () => {
    expect(parseSystemEvent('ThreadActivity/TopicUpdate', '<topicupdate><eventtime>1</eventtime><value>Q&amp;A sync</value></topicupdate>'))
      .toEqual({ type: 'topic_changed', summary: 'Chat renamed to "Q&A sync"' });
  });

  it('summarizes members joining from JSON, naming them only when all names are known', () => {
    const anonymous = JSON.stringify({ members: [{ id: '8:orgid:a', friendlyname: '' }, { id: '8:orgid:b', friendlyname: '' }] });
    expect(parseSystemEvent('ThreadActivity/MemberJoined', anonymous)).toEqual({ type: 'members_joined', summary: '2 members joined' });
    const named = JSON.stringify({ members: [{ id: '8:orgid:a', friendlyname: 'Ada Lovelace' }] });
    expect(parseSystemEvent('ThreadActivity/MemberJoined', named)).toEqual({ type: 'members_joined', summary: '1 member joined: Ada Lovelace' });
  });

  it('counts XML membership targets', () => {
    expect(parseSystemEvent('ThreadActivity/DeleteMember', '<deletemember><target>8:orgid:a</target><target>8:orgid:b</target></deletemember>'))
      .toEqual({ type: 'members_removed', summary: '2 members removed' });
  });

  it('does not invent a member count it cannot read', () => {
    expect(parseSystemEvent('ThreadActivity/MemberLeft', '<memberleft/>')).toEqual({ type: 'members_left', summary: 'Members left' });
  });

  it('labels unknown events without exposing their payload', () => {
    expect(parseSystemEvent('ThreadActivity/HistoryDisclosedUpdate', '<historydisclosedupdate><value>true</value></historydisclosedupdate>'))
      .toEqual({ type: 'history_disclosed_update', summary: 'Teams event: ThreadActivity/HistoryDisclosedUpdate' });
  });
});

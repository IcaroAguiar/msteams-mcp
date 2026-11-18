import { afterAll, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ErrorCode } from '../types/errors.js';
import { buildAmsImageHtml, buildFileEntry, prepareAttachments, readImageSize } from './upload-api.js';

function png(width: number, height: number): Buffer {
  const b = Buffer.alloc(24);
  b.writeUInt32BE(0x89504e47, 0);
  b.writeUInt32BE(width, 16);
  b.writeUInt32BE(height, 20);
  return b;
}

function jpeg(width: number, height: number): Buffer {
  const app0 = Buffer.from([0xff, 0xe0, 0x00, 0x10, ...Buffer.alloc(14)]);
  const sof = Buffer.alloc(19);
  sof.writeUInt16BE(0xffc2, 0);
  sof.writeUInt16BE(17, 2);
  sof.writeUInt16BE(height, 5);
  sof.writeUInt16BE(width, 7);
  return Buffer.concat([Buffer.from([0xff, 0xd8]), app0, sof]);
}

function gif(width: number, height: number): Buffer {
  const b = Buffer.alloc(10);
  b.write('GIF89a', 0, 'ascii');
  b.writeUInt16LE(width, 6);
  b.writeUInt16LE(height, 8);
  return b;
}

function webpExtended(width: number, height: number): Buffer {
  const b = Buffer.alloc(30);
  b.write('RIFF', 0, 'ascii');
  b.write('WEBPVP8X', 8, 'ascii');
  b.writeUIntLE(width - 1, 24, 3);
  b.writeUIntLE(height - 1, 27, 3);
  return b;
}

describe('readImageSize', () => {
  it.each([
    ['png', png(1440, 900)],
    ['progressive jpeg after an APP0 segment', jpeg(1440, 900)],
    ['gif', gif(1440, 900)],
    ['extended webp', webpExtended(1440, 900)],
  ])('reads %s dimensions', (_name, bytes) => {
    expect(readImageSize(bytes)).toEqual({ width: 1440, height: 900 });
  });

  it('returns null for bytes that are not an image', () => {
    expect(readImageSize(Buffer.from('%PDF-1.7 not an image'))).toBeNull();
  });
});

describe('message markup', () => {
  it('embeds an AMS image the way Teams renders a pasted picture', () => {
    const html = buildAmsImageHtml('https://us-api.asm.skype.com', '0-eus-d1-abc', { width: 10, height: 20 });
    expect(html).toContain('itemtype="http://schema.skype.com/AMSImage"');
    expect(html).toContain('src="https://us-api.asm.skype.com/v1/objects/0-eus-d1-abc/views/imgo"');
    expect(html).toContain('width="10" height="20"');
  });

  it('describes a OneDrive file with its unique id taken from the eTag', () => {
    const entry = buildFileEntry(
      { id: '01ABC', name: 'Relatorio.PDF', webUrl: 'https://x-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/Relatorio.PDF', eTag: '"{0B3B8F1C-1E2D-4F5A-9B8C-7D6E5F4A3B2C},1"' },
      'https://x-my.sharepoint.com/personal/u/'
    );
    expect(entry).toMatchObject({ id: '0B3B8F1C-1E2D-4F5A-9B8C-7D6E5F4A3B2C', itemid: '01ABC', fileType: 'pdf', title: 'Relatorio.PDF' });
  });
});

describe('prepareAttachments validation (before any upload)', () => {
  const dir = mkdtempSync(join(tmpdir(), 'teams-upload-'));
  const pdf = join(dir, 'doc.pdf');
  writeFileSync(pdf, '%PDF-1.7');
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  it('refuses a non-image file in a channel', async () => {
    const result = await prepareAttachments('19:abc@thread.tacv2', [pdf]);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.INVALID_INPUT);
      expect(result.error.message).toContain('Channels accept image attachments only');
    }
  });

  it('refuses relative and missing paths', async () => {
    for (const path of ['doc.pdf', join(dir, 'missing.png')]) {
      const result = await prepareAttachments('48:notes', [path]);
      expect(result.ok).toBe(false);
    }
  });
});

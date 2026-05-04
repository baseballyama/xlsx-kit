import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { commentsToBytes, parseCommentsXml, serializeComments } from '../../src/worksheet/comments-xml';
import { getComment, removeComment, setComment, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('parseCommentsXml + serializeComments', () => {
  it('round-trips a single comment with one author', () => {
    const xml = serializeComments([{ ref: 'A1', author: 'Alice', text: 'Hello' }]);
    const out = parseCommentsXml(xml);
    expect(out).toEqual([{ ref: 'A1', author: 'Alice', text: 'Hello' }]);
  });

  it('dedupes authors across multiple comments', () => {
    const xml = serializeComments([
      { ref: 'A1', author: 'Alice', text: 'one' },
      { ref: 'A2', author: 'Alice', text: 'two' },
      { ref: 'A3', author: 'Bob', text: 'three' },
    ]);
    expect(xml).toContain('<author>Alice</author>');
    expect(xml).toContain('<author>Bob</author>');
    // <author>Alice</author> appears exactly once.
    const aliceMatches = xml.match(/<author>Alice<\/author>/g);
    expect(aliceMatches?.length).toBe(1);
    const out = parseCommentsXml(xml);
    expect(out.map((c) => c.author)).toEqual(['Alice', 'Alice', 'Bob']);
  });

  it('escapes < > & " in text and author', () => {
    const xml = serializeComments([{ ref: 'A1', author: 'A & B', text: 'note < here >' }]);
    expect(xml).toContain('A &amp; B');
    expect(xml).toContain('note &lt; here &gt;');
    const out = parseCommentsXml(xml);
    expect(out[0]?.author).toBe('A & B');
    expect(out[0]?.text).toBe('note < here >');
  });

  it('rejects a non-comments root', () => {
    expect(() => parseCommentsXml('<foo/>')).toThrowError(/expected comments/);
  });

  it('handles <r>/<t> rich-text by concatenating run text', () => {
    const xml =
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>foo</t></r><r><t>bar</t></r></text></comment></commentList></comments>';
    expect(parseCommentsXml(xml)[0]?.text).toBe('foobar');
  });
});

describe('setComment / getComment / removeComment', () => {
  it('add, replace, get, remove', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setComment(ws, { ref: 'A1', author: 'Me', text: 'first' });
    expect(getComment(ws, 'A1')?.text).toBe('first');
    setComment(ws, { ref: 'A1', author: 'Me', text: 'second' });
    expect(ws.legacyComments.length).toBe(1);
    expect(getComment(ws, 'A1')?.text).toBe('second');
    expect(removeComment(ws, 'A1')).toBe(true);
    expect(removeComment(ws, 'A1')).toBe(false);
  });
});

describe('comments round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves a single sheet of comments + author dedup', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'hi' });
    setComment(ws, { ref: 'B2', author: 'Alice', text: 'check this' });
    setComment(ws, { ref: 'C3', author: 'Bob', text: 'fix me' });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.legacyComments.map((c) => c.ref).sort()).toEqual(['A1', 'B2', 'C3']);
    expect(getComment(ws2, 'A1')?.author).toBe('Alice');
    expect(getComment(ws2, 'C3')?.text).toBe('fix me');
  });

  it('preserves comments across multiple sheets with workbook-global commentsN ids', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setComment(a, { ref: 'A1', author: 'X', text: 'on A' });
    setComment(b, { ref: 'B2', author: 'Y', text: 'on B' });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const a2 = expectSheet(wb2.sheets[0]?.sheet);
    const b2 = expectSheet(wb2.sheets[1]?.sheet);
    expect(a2.legacyComments[0]?.text).toBe('on A');
    expect(b2.legacyComments[0]?.text).toBe('on B');
  });

  it('omits comments parts when no comments exist', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoComments');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('xl/comments');
    expect(txt).not.toContain('vmlDrawing');
    expect(txt).not.toContain('<legacyDrawing');
  });
});

describe('commentsToBytes', () => {
  it('returns Uint8Array bytes that parse back', () => {
    const bytes = commentsToBytes([{ ref: 'A1', author: 'A', text: 'hi' }]);
    expect(bytes).toBeInstanceOf(Uint8Array);
    const out = parseCommentsXml(bytes);
    expect(out[0]?.text).toBe('hi');
  });
});

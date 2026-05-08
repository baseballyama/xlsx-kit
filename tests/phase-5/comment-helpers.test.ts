// Tests for listComments / renameCommentAuthor / findCommentsByAuthor.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  findCommentsByAuthor,
  listComments,
  renameCommentAuthor,
  setComment,
  type Worksheet,
} from '../../src/xlsx/worksheet/worksheet';

describe('listComments', () => {
  it('returns the worksheet legacyComments array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'one' });
    setComment(ws, { ref: 'B2', author: 'Bob', text: 'two' });
    expect(listComments(ws).map((c) => c.author)).toEqual(['Alice', 'Bob']);
  });

  it('reflects subsequent removals', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'one' });
    setComment(ws, { ref: 'B2', author: 'Bob', text: 'two' });
    expect(listComments(ws).length).toBe(2);
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'updated' });
    expect(listComments(ws).length).toBe(2);
    expect(listComments(ws)[0]?.text).toBe('updated');
  });
});

describe('renameCommentAuthor', () => {
  it('returns the count and rewrites every match', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Old', text: 'a' });
    setComment(ws, { ref: 'A2', author: 'Other', text: 'x' });
    setComment(ws, { ref: 'A3', author: 'Old', text: 'c' });
    expect(renameCommentAuthor(ws, 'Old', 'New')).toBe(2);
    expect(listComments(ws).map((c) => c.author).sort()).toEqual(['New', 'New', 'Other']);
  });

  it('returns 0 when no comment matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'x' });
    expect(renameCommentAuthor(ws, 'Bob', 'Bobby')).toBe(0);
    expect(listComments(ws)[0]?.author).toBe('Alice');
  });

  it('round-trip after rename: the new name is what comes back', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Old', text: 'one' });
    setComment(ws, { ref: 'B2', author: 'Old', text: 'two' });
    renameCommentAuthor(ws, 'Old', 'New');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.legacyComments.every((c) => c.author === 'New')).toBe(true);
  });
});

describe('findCommentsByAuthor', () => {
  it('filters by author', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'a1' });
    setComment(ws, { ref: 'B1', author: 'Bob', text: 'b1' });
    setComment(ws, { ref: 'A2', author: 'Alice', text: 'a2' });
    const aliceOnly = findCommentsByAuthor(ws, 'Alice');
    expect(aliceOnly.map((c) => c.ref).sort()).toEqual(['A1', 'A2']);
  });

  it('returns empty array when nothing matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'x' });
    expect(findCommentsByAuthor(ws, 'Carol')).toEqual([]);
  });
});

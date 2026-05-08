// Tests for removeAllHyperlinks / removeAllComments.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { addUrlHyperlink } from '../../src/xlsx/worksheet/hyperlinks';
import {
  removeAllComments,
  removeAllHyperlinks,
  setComment,
} from '../../src/xlsx/worksheet/worksheet';

describe('removeAllHyperlinks', () => {
  it('drops every hyperlink and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'A1', 'https://a.example');
    addUrlHyperlink(ws, 'B2', 'https://b.example');
    expect(removeAllHyperlinks(ws)).toBe(2);
    expect(ws.hyperlinks).toEqual([]);
  });

  it('returns 0 when no hyperlinks exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllHyperlinks(ws)).toBe(0);
  });
});

describe('removeAllComments', () => {
  it('drops every comment and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'a' });
    setComment(ws, { ref: 'B2', author: 'Bob', text: 'b' });
    expect(removeAllComments(ws)).toBe(2);
    expect(ws.legacyComments).toEqual([]);
  });

  it('returns 0 when no comments exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllComments(ws)).toBe(0);
  });
});

// Tests for getCellComment resolver.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getCellComment,
  setCell,
  setComment,
} from '../../src/xlsx/worksheet/worksheet';

describe('getCellComment', () => {
  it('resolves a single-cell ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'note' });
    expect(getCellComment(ws, c)?.text).toBe('note');
  });

  it('returns undefined when no comment matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setComment(ws, { ref: 'B2', author: 'Alice', text: 'elsewhere' });
    expect(getCellComment(ws, c)).toBeUndefined();
  });

  it('multi-comment sheet returns the matching one only', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'a' });
    setComment(ws, { ref: 'C3', author: 'Bob', text: 'c' });
    const c1 = setCell(ws, 1, 1, '1');
    const c3 = setCell(ws, 3, 3, '3');
    expect(getCellComment(ws, c1)?.text).toBe('a');
    expect(getCellComment(ws, c3)?.text).toBe('c');
  });

  it('range-style ref containment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'B2:D4', author: 'Alice', text: 'rangewide' });
    const inside = setCell(ws, 3, 3, 'mid');
    const outside = setCell(ws, 5, 5, 'out');
    expect(getCellComment(ws, inside)?.text).toBe('rangewide');
    expect(getCellComment(ws, outside)).toBeUndefined();
  });
});

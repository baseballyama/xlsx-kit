// Tests for getCellHyperlink resolver.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  addInternalHyperlink,
  addUrlHyperlink,
} from '../../src/xlsx/worksheet/hyperlinks';
import {
  getCellHyperlink,
  setCell,
} from '../../src/xlsx/worksheet/worksheet';

describe('getCellHyperlink', () => {
  it('resolves a single-cell ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'home');
    addUrlHyperlink(ws, 'A1', 'https://example.com');
    expect(getCellHyperlink(ws, c)?.target).toBe('https://example.com');
  });

  it('resolves a range ref by containment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'B2:D4', 'https://example.com');
    const inside = setCell(ws, 3, 3, 'mid');
    const outside = setCell(ws, 5, 5, 'out');
    expect(getCellHyperlink(ws, inside)?.target).toBe('https://example.com');
    expect(getCellHyperlink(ws, outside)).toBeUndefined();
  });

  it('returns undefined when there are no hyperlinks', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'plain');
    expect(getCellHyperlink(ws, c)).toBeUndefined();
  });

  it('first matching entry wins (insertion order)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'A1:B2', 'https://outer.example.com');
    addUrlHyperlink(ws, 'A1', 'https://inner.example.com');
    const c = setCell(ws, 1, 1, 'shared');
    expect(getCellHyperlink(ws, c)?.target).toBe('https://outer.example.com');
  });

  it('internal hyperlink is also resolvable', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addInternalHyperlink(ws, 'A1', "'Sheet 2'!A1");
    const c = setCell(ws, 1, 1, 'jump');
    expect(getCellHyperlink(ws, c)?.location).toBe("'Sheet 2'!A1");
  });
});

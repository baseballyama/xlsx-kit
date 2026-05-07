// Tests for renameColumns — bulk header rename.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  readRangeAsObjects,
  renameColumns,
  setCell,
} from '../../src/worksheet/worksheet';

describe('renameColumns', () => {
  it('renames a single column via a one-key mapping', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    renameColumns(ws, 'A1:A2', { name: 'fullName' });
    expect(readRangeAsObjects(ws, 'A1:A2')).toEqual([{ fullName: 'Alice' }]);
  });

  it('applies multiple renames in one call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b'); setCell(ws, 1, 3, 'c');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2); setCell(ws, 2, 3, 3);
    renameColumns(ws, 'A1:C2', { a: 'alpha', c: 'gamma' });
    expect(readRangeAsObjects(ws, 'A1:C2')).toEqual([{ alpha: 1, b: 2, gamma: 3 }]);
  });

  it('supports a column-swap (a→b and b→a) without collision errors', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2);
    renameColumns(ws, 'A1:B2', { a: 'b', b: 'a' });
    // Header positions are unchanged (just relabelled), so col 1 = 'b', col 2 = 'a'.
    expect(readRangeAsObjects(ws, 'A1:B2')).toEqual([{ b: 1, a: 2 }]);
  });

  it('throws when an oldName is not in the header row (no partial mutation)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    expect(() => renameColumns(ws, 'A1:B1', { a: 'alpha', missing: 'x' })).toThrow(/missing/);
    // First entry must NOT have been applied since validation runs up-front.
    expect(ws.rows.get(1)?.get(1)?.value).toBe('a');
  });

  it('throws when newName collides with another existing header (not being renamed away)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    expect(() => renameColumns(ws, 'A1:B1', { a: 'b' })).toThrow(/already exists/);
  });

  it('throws when two entries map to the same newName', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b'); setCell(ws, 1, 3, 'c');
    expect(() => renameColumns(ws, 'A1:C1', { a: 'x', b: 'x' })).toThrow(/duplicate target/);
  });

  it('empty mapping is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    renameColumns(ws, 'A1:A1', {});
    expect(ws.rows.get(1)?.get(1)?.value).toBe('a');
  });
});

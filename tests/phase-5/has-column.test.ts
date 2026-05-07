// Tests for hasColumn — header includes predicate.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { hasColumn, setCell } from '../../src/worksheet/worksheet';

describe('hasColumn', () => {
  it('returns true when the column exists in the header', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    expect(hasColumn(ws, 'A1:B1', 'age')).toBe(true);
  });

  it('returns false when the column is not in the header', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    expect(hasColumn(ws, 'A1:A1', 'role')).toBe(false);
  });

  it('returns false for an empty range header', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(hasColumn(ws, 'A1:C1', 'anything')).toBe(false);
  });

  it('matches the empty-string sentinel for unmaterialised header cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    // col 2 left empty → '' header sentinel
    setCell(ws, 1, 3, 'c');
    expect(hasColumn(ws, 'A1:C1', '')).toBe(true);
  });
});

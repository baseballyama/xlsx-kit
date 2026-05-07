// Tests for getValueAtAddress — sheet-qualified A1 → CellValue | null.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getValueAtAddress,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getValueAtAddress', () => {
  it('returns the cell value when materialised', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, 42);
    expect(getValueAtAddress(wb, 'Data!A1')).toBe('a');
    expect(getValueAtAddress(wb, 'Data!A2')).toBe(42);
  });

  it('returns null for unmaterialised cells', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(getValueAtAddress(wb, 'Data!Z99')).toBeNull();
  });

  it('honours quoted sheet titles', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 1, 1, 'q');
    expect(getValueAtAddress(wb, "'Q1 2024'!A1")).toBe('q');
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => getValueAtAddress(wb, 'Missing!A1')).toThrow(/sheet/);
  });
});

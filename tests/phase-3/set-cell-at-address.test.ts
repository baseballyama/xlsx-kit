// Tests for setCellAtAddress — sheet-qualified A1 → single-cell write.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getCellAtAddress,
  setCellAtAddress,
} from '../../src/workbook/workbook';

describe('setCellAtAddress', () => {
  it('writes a value at the address and returns the Cell', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    const cell = setCellAtAddress(wb, 'Data!B2', 'hello');
    expect(cell.row).toBe(2);
    expect(cell.col).toBe(2);
    expect(cell.value).toBe('hello');
    expect(getCellAtAddress(wb, 'Data!B2')?.value).toBe('hello');
  });

  it('honours quoted sheet titles', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Q1 2024');
    setCellAtAddress(wb, "'Q1 2024'!A1", 42);
    expect(getCellAtAddress(wb, "'Q1 2024'!A1")?.value).toBe(42);
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => setCellAtAddress(wb, 'Missing!A1', 'x')).toThrow(/sheet/);
  });

  it('throws when the address points at a range', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => setCellAtAddress(wb, 'Data!A1:B5', 'x')).toThrow(/range/);
  });
});

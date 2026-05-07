// Tests for getCellAtAddress — sheet-qualified A1 → Cell lookup.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, getCellAtAddress } from '../../src/workbook/workbook';
import { getCellAddress, setCell } from '../../src/worksheet/worksheet';

describe('getCellAtAddress', () => {
  it('resolves a bare-title address', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    expect(getCellAtAddress(wb, 'Data!A1')?.value).toBe('a');
  });

  it('resolves a quoted-title address', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 5, 7, 'cell');
    expect(getCellAtAddress(wb, "'Q1 2024'!G5")?.value).toBe('cell');
  });

  it('returns undefined when the cell is not materialised', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(getCellAtAddress(wb, 'Data!Z99')).toBeUndefined();
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => getCellAtAddress(wb, 'Missing!A1')).toThrow(/sheet/);
  });

  it('throws when the address points at a range instead of a single cell', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => getCellAtAddress(wb, 'Data!A1:B5')).toThrow(/range/);
  });

  it('round-trips through getCellAddress', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 3, 4, 'rt');
    const cell = ws.rows.get(3)?.get(4);
    if (!cell) throw new Error('cell missing');
    const address = getCellAddress(ws, cell);
    expect(getCellAtAddress(wb, address)?.value).toBe('rt');
  });
});

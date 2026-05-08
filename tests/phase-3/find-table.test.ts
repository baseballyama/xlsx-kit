// Tests for workbook-wide findTable.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  findTable,
} from '../../src/xlsx/workbook/workbook';
import { addExcelTable } from '../../src/xlsx/worksheet/table';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('findTable', () => {
  it('locates a table by displayName across multiple sheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'h');
    setCell(b, 1, 1, 'h');
    addExcelTable(wb, a, { name: 'TblA', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addExcelTable(wb, b, { name: 'TblB', ref: 'A1:B2', columns: ['c1', 'c2'] });
    const hit = findTable(wb, 'TblB');
    expect(hit?.sheet.title).toBe('B');
    expect(hit?.table.name).toBe('TblB');
  });

  it('returns undefined when the table is not registered', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'h');
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });
    expect(findTable(wb, 'Missing')).toBeUndefined();
  });

  it('empty workbook → undefined', () => {
    const wb = createWorkbook();
    expect(findTable(wb, 'Anything')).toBeUndefined();
  });
});

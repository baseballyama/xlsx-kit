// Tests for the defined-name builder helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  addDefinedName,
  getDefinedName,
  removeDefinedName,
  setPrintArea,
  setPrintTitles,
} from '../../src/workbook/defined-names';

describe('addDefinedName / getDefinedName / removeDefinedName', () => {
  it('addDefinedName replaces an existing entry with the same name + scope', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addDefinedName(wb, { name: 'tax', value: '0.05' });
    expect(wb.definedNames).toHaveLength(1);
    addDefinedName(wb, { name: 'tax', value: '0.08' });
    expect(wb.definedNames).toHaveLength(1);
    expect(wb.definedNames[0]?.value).toBe('0.08');
  });

  it('different scopes coexist (workbook + sheet 0)', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addDefinedName(wb, { name: 'region', value: '"All"' });
    addDefinedName(wb, { name: 'region', value: '"North"', scope: 0 });
    expect(wb.definedNames).toHaveLength(2);
    expect(getDefinedName(wb, 'region')?.value).toBe('"All"');
    expect(getDefinedName(wb, 'region', 0)?.value).toBe('"North"');
  });

  it('removeDefinedName returns true when removed, false otherwise', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'x', value: '1' });
    expect(removeDefinedName(wb, 'x')).toBe(true);
    expect(removeDefinedName(wb, 'x')).toBe(false);
    expect(wb.definedNames).toHaveLength(0);
  });
});

describe('setPrintArea / setPrintTitles', () => {
  it('setPrintArea writes _xlnm.Print_Area scoped to the sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    const dn = setPrintArea(wb, 0, 'A1:E20');
    expect(dn.name).toBe('_xlnm.Print_Area');
    expect(dn.value).toBe('A1:E20');
    expect(dn.scope).toBe(0);
  });

  it('setPrintTitles formats both rows + cols with sheet prefix', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    const dn = setPrintTitles(wb, 0, { rows: '$1:$1', cols: '$A:$A', sheetName: 'Report' });
    expect(dn.value).toBe("'Report'!$A:$A,'Report'!$1:$1");
  });

  it('setPrintTitles supports rows-only or cols-only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const dnRows = setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: 'A' });
    expect(dnRows.value).toBe("'A'!$1:$1");
    const dnCols = setPrintTitles(wb, 0, { cols: '$A:$A', sheetName: 'A' });
    expect(dnCols.value).toBe("'A'!$A:$A");
  });

  it('setPrintTitles throws when neither rows nor cols is supplied', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => setPrintTitles(wb, 0, { sheetName: 'A' })).toThrow();
  });

  it('full save → load round-trip preserves the print-area + titles', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setPrintArea(wb, 0, 'A1:E20');
    setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: 'A' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const pa = getDefinedName(wb2, '_xlnm.Print_Area', 0);
    const pt = getDefinedName(wb2, '_xlnm.Print_Titles', 0);
    expect(pa?.value).toBe('A1:E20');
    expect(pt?.value).toBe("'A'!$1:$1");
  });
});
// Tests for listPrintAreas / listPrintTitles.

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  listPrintAreas,
  listPrintTitles,
  setPrintArea,
  setPrintTitles,
} from '../../src/xlsx/workbook/defined-names';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('listPrintAreas', () => {
  it('returns only _xlnm.Print_Area entries', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    setPrintArea(wb, 0, "'A'!$A$1:$D$10");
    setPrintArea(wb, 1, "'B'!$A$1:$E$20");
    addDefinedName(wb, { name: 'Other', value: '$A$1' });
    const out = listPrintAreas(wb);
    expect(out.length).toBe(2);
    expect(out.every((d) => d.name === '_xlnm.Print_Area')).toBe(true);
  });

  it('empty when no print areas defined', () => {
    const wb = createWorkbook();
    expect(listPrintAreas(wb)).toEqual([]);
  });
});

describe('listPrintTitles', () => {
  it('returns only _xlnm.Print_Titles entries', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: 'A' });
    addDefinedName(wb, { name: 'Other', value: '$A$1' });
    const out = listPrintTitles(wb);
    expect(out.length).toBe(1);
    expect(out[0]?.name).toBe('_xlnm.Print_Titles');
    expect(out[0]?.value).toContain('$1:$1');
  });
});

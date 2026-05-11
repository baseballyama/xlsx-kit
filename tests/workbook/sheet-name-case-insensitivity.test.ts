// Excel treats sheet names as case-insensitive for uniqueness — "Data" and
// "data" cannot co-exist in the same workbook. These tests pin that behavior
// down for addWorksheet / addChartsheet / pickUniqueSheetTitle so regressions
// don't quietly produce workbooks Excel refuses to open.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  duplicateSheet,
  pickUniqueSheetTitle,
} from '../../src/workbook/workbook';

describe('sheet name uniqueness — case-insensitive', () => {
  it('addWorksheet rejects a title that differs only in case', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => addWorksheet(wb, 'data')).toThrow(/already in use/);
    expect(() => addWorksheet(wb, 'DATA')).toThrow(/already in use/);
    expect(() => addWorksheet(wb, 'DaTa')).toThrow(/already in use/);
  });

  it('addChartsheet collides with a worksheet of the same name in a different case', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Charts');
    expect(() => addChartsheet(wb, 'charts')).toThrow(/already in use/);
  });

  it('addWorksheet collides with a chartsheet of the same name in a different case', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Charts');
    expect(() => addWorksheet(wb, 'charts')).toThrow(/already in use/);
  });

  it('pickUniqueSheetTitle picks the next slot when only a case variant exists', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(pickUniqueSheetTitle(wb, 'DATA')).toBe('DATA (2)');
    expect(pickUniqueSheetTitle(wb, 'data')).toBe('data (2)');
  });

  it('duplicateSheet rejects a target that differs only in case', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => duplicateSheet(wb, 'Data', 'data')).toThrow(/already in use/);
  });
});

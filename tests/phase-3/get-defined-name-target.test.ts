// Tests for getDefinedNameTarget — DefinedName.value → parsed legs.

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  addDefinedNameForRange,
  getDefinedNameTarget,
} from '../../src/xlsx/workbook/defined-names';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('getDefinedNameTarget', () => {
  it('parses a single Sheet!A1:B5 leg', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    addDefinedNameForRange(wb, 'X', ws, 'A1:B5');
    const targets = getDefinedNameTarget(wb, 'X');
    expect(targets).toEqual([
      { sheet: 'Data', range: 'A1:B5', bounds: { minRow: 1, minCol: 1, maxRow: 5, maxCol: 2 } },
    ]);
  });

  it('handles a quoted sheet title', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    addDefinedNameForRange(wb, 'Y', ws, 'C2:C100');
    const targets = getDefinedNameTarget(wb, 'Y');
    expect(targets?.[0]?.sheet).toBe('Q1 2024');
    expect(targets?.[0]?.range).toBe('C2:C100');
  });

  it('splits comma-separated multi-range Print_Titles-style values', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    addDefinedName(wb, {
      name: '_xlnm.Print_Titles',
      value: "'Data'!$1:$1,'Data'!$A:$A",
      scope: 0,
    });
    const targets = getDefinedNameTarget(wb, '_xlnm.Print_Titles', 0);
    expect(targets?.length).toBe(2);
    expect(targets?.[0]?.sheet).toBe('Data');
    expect(targets?.[0]?.range).toBe('$1:$1');
    expect(targets?.[1]?.range).toBe('$A:$A');
  });

  it('returns undefined when the defined name does not exist', () => {
    const wb = createWorkbook();
    expect(getDefinedNameTarget(wb, 'Missing')).toBeUndefined();
  });
});

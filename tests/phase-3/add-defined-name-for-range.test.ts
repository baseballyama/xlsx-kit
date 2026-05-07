// Tests for addDefinedNameForRange — sheet-qualified DefinedName builder.

import { describe, expect, it } from 'vitest';
import { addDefinedNameForRange } from '../../src/workbook/defined-names';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('addDefinedNameForRange', () => {
  it('registers a workbook-scoped name with the sheet-qualified ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const dn = addDefinedNameForRange(wb, 'MyRange', ws, 'A1:B5');
    expect(dn.name).toBe('MyRange');
    expect(dn.value).toBe('Data!A1:B5');
    expect(dn.scope).toBeUndefined();
    expect(wb.definedNames.length).toBe(1);
  });

  it('quotes the sheet name when needed', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    const dn = addDefinedNameForRange(wb, 'Sales', ws, 'C2:C100');
    expect(dn.value).toBe("'Q1 2024'!C2:C100");
  });

  it('opts.localToSheet scopes the name to the worksheet index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'First');
    const ws = addWorksheet(wb, 'Second');
    const dn = addDefinedNameForRange(wb, 'Local', ws, 'A1', { localToSheet: true });
    expect(dn.scope).toBe(1);
  });

  it('throws when localToSheet is set but the worksheet is not on the workbook', () => {
    const wb = createWorkbook();
    const otherWb = createWorkbook();
    const stale = addWorksheet(otherWb, 'Stale');
    expect(() =>
      addDefinedNameForRange(wb, 'X', stale, 'A1', { localToSheet: true }),
    ).toThrow(/not registered/);
  });

  it('replacing an existing name with same scope updates in-place (no duplicate row)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    addDefinedNameForRange(wb, 'X', ws, 'A1');
    const updated = addDefinedNameForRange(wb, 'X', ws, 'B2');
    expect(wb.definedNames.length).toBe(1);
    expect(updated.value).toBe('Data!B2');
  });
});

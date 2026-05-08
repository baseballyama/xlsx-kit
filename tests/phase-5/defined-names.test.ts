import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { makeDefinedName } from '../../src/xlsx/workbook/defined-names';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('makeDefinedName', () => {
  it('honours name + value, drops absent optionals', () => {
    const dn = makeDefinedName({ name: 'MyRange', value: "'Sheet 1'!$A$1:$B$10" });
    expect(dn.name).toBe('MyRange');
    expect(dn.value).toBe("'Sheet 1'!$A$1:$B$10");
    expect(dn.scope).toBeUndefined();
    expect(dn.hidden).toBeUndefined();
    expect(dn.comment).toBeUndefined();
  });

  it('records every optional when set', () => {
    const dn = makeDefinedName({
      name: 'Hidden',
      value: 'Sheet1!$A$1',
      scope: 0,
      hidden: true,
      comment: 'internal use',
    });
    expect(dn).toEqual({
      name: 'Hidden',
      value: 'Sheet1!$A$1',
      scope: 0,
      hidden: true,
      comment: 'internal use',
    });
  });
});

describe('definedNames round-trip through saveWorkbook → loadWorkbook', () => {
  it('omits the <definedNames> block when none are set', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    const bytes = await workbookToBytes(wb);
    // Workbook XML should not carry the wrapper.
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('<definedNames');
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.definedNames).toEqual([]);
  });

  it('preserves a workbook-scope name', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    wb.definedNames.push(makeDefinedName({ name: 'TaxRate', value: '0.08' }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.definedNames).toEqual([{ name: 'TaxRate', value: '0.08' }]);
  });

  it('preserves a sheet-scope name with localSheetId', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    addWorksheet(wb, 'Sheet2');
    wb.definedNames.push(makeDefinedName({ name: 'Sheet2Range', value: 'Sheet2!$A$1:$Z$100', scope: 1 }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.definedNames[0]?.scope).toBe(1);
    expect(wb2.definedNames[0]?.value).toBe('Sheet2!$A$1:$Z$100');
  });

  it('preserves the _xlnm.Print_Area built-in name', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    wb.definedNames.push(makeDefinedName({ name: '_xlnm.Print_Area', value: 'Report!$A$1:$E$50', scope: 0 }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.definedNames[0]?.name).toBe('_xlnm.Print_Area');
    expect(wb2.definedNames[0]?.scope).toBe(0);
  });

  it('preserves hidden + comment + escapes special chars in value', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    wb.definedNames.push(
      makeDefinedName({
        name: 'WithChars',
        value: 'IF(A1<10,"low & slow","fast")',
        hidden: true,
        comment: 'note < & >',
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const dn = wb2.definedNames[0];
    expect(dn?.value).toBe('IF(A1<10,"low & slow","fast")');
    expect(dn?.hidden).toBe(true);
    expect(dn?.comment).toBe('note < & >');
  });

  it('preserves order across multiple names', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    wb.definedNames.push(makeDefinedName({ name: 'First', value: 'A!$A$1' }));
    wb.definedNames.push(makeDefinedName({ name: 'Second', value: 'B!$A$1', scope: 1 }));
    wb.definedNames.push(makeDefinedName({ name: 'Third', value: 'A!$Z$99' }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.definedNames.map((d) => d.name)).toEqual(['First', 'Second', 'Third']);
  });
});

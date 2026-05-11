// Tests for the typed Workbook.externalReferences link array.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('externalReferences round-trip', () => {
  it('preserves the rId list', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.externalReferences = [{ rId: 'rIdExt1' }, { rId: 'rIdExt2' }];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.externalReferences?.length).toBe(2);
    expect(wb2.externalReferences?.[0]?.rId).toBe('rIdExt1');
    expect(wb2.externalReferences?.[1]?.rId).toBe('rIdExt2');
  });

  it('emits no <externalReferences/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.externalReferences).toBeUndefined();
  });

  it('emits no element when the array is empty', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.externalReferences = [];
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.externalReferences).toBeUndefined();
  });
});
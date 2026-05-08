// Tests for the typed workbook-level <oleSize> and <fileRecoveryPr>.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeFileRecoveryProperties } from '../../src/xlsx/workbook/file-recovery';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('oleSize round-trip', () => {
  it('preserves the bounding ref', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.oleSize = 'A1:E20';

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.oleSize).toBe('A1:E20');
  });

  it('emits no <oleSize/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.oleSize).toBeUndefined();
  });
});

describe('fileRecoveryPr round-trip', () => {
  it('preserves the four boolean flags', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'R');
    wb.fileRecoveryPr = makeFileRecoveryProperties({
      autoRecover: false,
      crashSave: true,
      dataExtractLoad: false,
      repairLoad: true,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const fp = wb2.fileRecoveryPr;
    expect(fp?.autoRecover).toBe(false);
    expect(fp?.crashSave).toBe(true);
    expect(fp?.dataExtractLoad).toBe(false);
    expect(fp?.repairLoad).toBe(true);
  });

  it('emits no <fileRecoveryPr/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.fileRecoveryPr).toBeUndefined();
  });
});
// Tests for the typed workbook-level <workbookProtection> model.
// Per docs/plan/13-full-excel-coverage.md §B5 (workbook side).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeWorkbookProtection } from '../../src/xlsx/workbook/protection';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('workbookProtection round-trip', () => {
  it('preserves the 3 lock flags', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'L');
    setCell(ws, 1, 1, 1);
    wb.workbookProtection = makeWorkbookProtection({
      lockStructure: true,
      lockWindows: false,
      lockRevision: true,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const wp = wb2.workbookProtection;
    expect(wp).toBeDefined();
    expect(wp?.lockStructure).toBe(true);
    expect(wp?.lockWindows).toBe(false);
    expect(wp?.lockRevision).toBe(true);
  });

  it('round-trips both modern hash quads (workbook + revisions)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    wb.workbookProtection = makeWorkbookProtection({
      lockStructure: true,
      workbookAlgorithmName: 'SHA-512',
      workbookHashValue: 'abc==',
      workbookSaltValue: 'def==',
      workbookSpinCount: 100000,
      revisionsAlgorithmName: 'SHA-512',
      revisionsHashValue: 'ghi==',
      revisionsSaltValue: 'jkl==',
      revisionsSpinCount: 200000,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const wp = wb2.workbookProtection;
    expect(wp?.workbookAlgorithmName).toBe('SHA-512');
    expect(wp?.workbookHashValue).toBe('abc==');
    expect(wp?.workbookSaltValue).toBe('def==');
    expect(wp?.workbookSpinCount).toBe(100000);
    expect(wp?.revisionsAlgorithmName).toBe('SHA-512');
    expect(wp?.revisionsHashValue).toBe('ghi==');
    expect(wp?.revisionsSaltValue).toBe('jkl==');
    expect(wp?.revisionsSpinCount).toBe(200000);
  });

  it('round-trips the legacy 16-bit hex hashes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'L');
    setCell(ws, 1, 1, 1);
    wb.workbookProtection = makeWorkbookProtection({
      workbookPassword: 'CC1A',
      revisionsPassword: 'D7F0',
      lockStructure: true,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProtection?.workbookPassword).toBe('CC1A');
    expect(wb2.workbookProtection?.revisionsPassword).toBe('D7F0');
  });

  it('emits no <workbookProtection> when undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProtection).toBeUndefined();
  });
});
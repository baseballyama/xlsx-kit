// Tests for the typed workbook-level <workbookPr> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeWorkbookProperties } from '../../src/workbook/workbook-properties';
import { setCell } from '../../src/worksheet/worksheet';

describe('workbookPr round-trip', () => {
  it('preserves codeName + defaultThemeVersion + updateLinks + showObjects', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.workbookProperties = makeWorkbookProperties({
      codeName: 'ThisWorkbook',
      defaultThemeVersion: 153222,
      updateLinks: 'never',
      showObjects: 'placeholders',
      hidePivotFieldList: true,
      filterPrivacy: false,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const wp = wb2.workbookProperties;
    expect(wp?.codeName).toBe('ThisWorkbook');
    expect(wp?.defaultThemeVersion).toBe(153222);
    expect(wp?.updateLinks).toBe('never');
    expect(wp?.showObjects).toBe('placeholders');
    expect(wp?.hidePivotFieldList).toBe(true);
    expect(wp?.filterPrivacy).toBe(false);
  });

  it('round-trips date1904 via the typed model', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.workbookProperties = makeWorkbookProperties({
      date1904: true,
      backupFile: false,
      saveExternalLinkValues: true,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProperties?.date1904).toBe(true);
    expect(wb2.workbookProperties?.backupFile).toBe(false);
    expect(wb2.workbookProperties?.saveExternalLinkValues).toBe(true);
    // Top-level wb.date1904 is what the cell-serial paths read; verify
    // it stays in sync via the existing parseDate1904 hook.
    expect(wb2.date1904).toBe(true);
  });

  it('synthesises a minimal workbookPr when only wb.date1904 is set', async () => {
    const wb = createWorkbook({ date1904: true });
    addWorksheet(wb, 'A');
    expect(wb.workbookProperties).toBeUndefined();

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    // The reader lifts <workbookPr date1904="1"/> back into the typed
    // model on the second load.
    expect(wb2.workbookProperties?.date1904).toBe(true);
    expect(wb2.date1904).toBe(true);
  });

  it('drops unknown enum values silently', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.workbookProperties = makeWorkbookProperties({
      showObjects: 'gibberish' as never,
      updateLinks: 'always',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProperties?.showObjects).toBeUndefined();
    expect(wb2.workbookProperties?.updateLinks).toBe('always');
  });

  it('emits no <workbookPr/> when nothing is set and date1904 is false', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProperties).toBeUndefined();
    expect(wb2.date1904).toBe(false);
  });
});
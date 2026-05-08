// Tests for the typed workbook-level <fileVersion> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeFileVersion } from '../../src/workbook/file-version';
import { setCell } from '../../src/worksheet/worksheet';

describe('fileVersion round-trip', () => {
  it('preserves all five attributes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.fileVersion = makeFileVersion({
      appName: 'xl',
      lastEdited: '7',
      lowestEdited: '7',
      rupBuild: '24827',
      codeName: '{F1A6E89C-1234-4567-89AB-CDEF01234567}',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const fv = wb2.fileVersion;
    expect(fv?.appName).toBe('xl');
    expect(fv?.lastEdited).toBe('7');
    expect(fv?.lowestEdited).toBe('7');
    expect(fv?.rupBuild).toBe('24827');
    expect(fv?.codeName).toBe('{F1A6E89C-1234-4567-89AB-CDEF01234567}');
  });

  it('emits no <fileVersion/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.fileVersion).toBeUndefined();
  });
});
// Tests for the typed workbook-level <fileSharing> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeFileSharing } from '../../src/workbook/file-sharing';
import { setCell } from '../../src/worksheet/worksheet';

describe('fileSharing round-trip', () => {
  it('preserves readOnlyRecommended + userName + modern hash quad', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.fileSharing = makeFileSharing({
      readOnlyRecommended: true,
      userName: 'qa',
      algorithmName: 'SHA-512',
      hashValue: 'aGFzaA==',
      saltValue: 'c2FsdA==',
      spinCount: 100000,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const fs = wb2.fileSharing;
    expect(fs?.readOnlyRecommended).toBe(true);
    expect(fs?.userName).toBe('qa');
    expect(fs?.algorithmName).toBe('SHA-512');
    expect(fs?.hashValue).toBe('aGFzaA==');
    expect(fs?.saltValue).toBe('c2FsdA==');
    expect(fs?.spinCount).toBe(100000);
  });

  it('preserves the legacy reservationPassword hex hash', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'B');
    wb.fileSharing = makeFileSharing({
      readOnlyRecommended: false,
      reservationPassword: 'CC1A',
      userName: 'admin',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.fileSharing?.reservationPassword).toBe('CC1A');
    expect(wb2.fileSharing?.userName).toBe('admin');
    expect(wb2.fileSharing?.readOnlyRecommended).toBe(false);
  });

  it('emits no <fileSharing/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.fileSharing).toBeUndefined();
  });
});
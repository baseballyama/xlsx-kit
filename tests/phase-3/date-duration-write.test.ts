// Phase 3 §5.5 — Date / Duration cell write.
// Covers what was deferred: setCell with a Date or {kind:'duration'}
// value now serialises through `dateToExcel` / `durationToExcel` into
// the workbook's epoch (Windows 1900 by default; Mac 1904 honoured).

import { describe, expect, it } from 'vitest';
import { makeDurationValue } from '../../src/xlsx/cell/cell';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { dateToExcel, durationToExcel } from '../../src/utils/datetime';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('phase-3 §5.5 — Date / Duration cell write', () => {
  it('writes a Date cell as the Windows-epoch serial number', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const date = new Date(Date.UTC(2026, 4, 5)); // 2026-05-05
    setCell(ws, 1, 1, date);
    const expectedSerial = dateToExcel(date, { epoch: 'windows' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.get(1)?.get(1)?.value).toBe(expectedSerial);
  });

  it('honours wb.date1904 when emitting Date serials', async () => {
    const wb = createWorkbook({ date1904: true });
    const ws = addWorksheet(wb, 'Sheet1');
    const date = new Date(Date.UTC(2026, 4, 5));
    setCell(ws, 1, 1, date);
    const expectedSerial = dateToExcel(date, { epoch: 'mac' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.get(1)?.get(1)?.value).toBe(expectedSerial);
  });

  it('writes a Duration cell as a fraction-of-day serial', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    // 90 minutes
    const ms = 90 * 60 * 1000;
    setCell(ws, 1, 1, makeDurationValue(ms));
    const expectedSerial = durationToExcel(ms);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.get(1)?.get(1)?.value).toBe(expectedSerial);
  });
});

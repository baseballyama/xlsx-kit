// Tests for the typed `<sheetPr>` model. (sheet view 拡張: tabColor / outlinePr /
// pageSetUpPr).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { makeColor } from '../../src/styles/colors';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeSheetProperties } from '../../src/worksheet/properties';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('sheetProperties API', () => {
  it('factory builds a typed shell with only the fields you set', () => {
    const sp = makeSheetProperties({
      codeName: 'Sheet1',
      tabColor: makeColor({ rgb: 'FF0070C0' }),
      outlinePr: { summaryBelow: false, summaryRight: false },
    });
    expect(sp.codeName).toBe('Sheet1');
    expect(sp.tabColor?.rgb).toBe('FF0070C0');
    expect(sp.outlinePr?.summaryBelow).toBe(false);
    expect(sp.outlinePr?.summaryRight).toBe(false);
    expect(sp.pageSetUpPr).toBeUndefined();
  });

  it('round-trips tabColor + codeName + outline + pageSetup through save/load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Demo');
    setCell(ws, 1, 1, 1);
    ws.sheetProperties = makeSheetProperties({
      codeName: 'DemoSheet',
      filterMode: false,
      published: true,
      tabColor: makeColor({ rgb: 'FF00B050' }),
      outlinePr: { summaryBelow: false, summaryRight: true, applyStyles: true },
      pageSetUpPr: { fitToPage: true, autoPageBreaks: false },
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const sp = ws2.sheetProperties;

    expect(sp).toBeDefined();
    expect(sp?.codeName).toBe('DemoSheet');
    expect(sp?.filterMode).toBe(false);
    expect(sp?.published).toBe(true);
    expect(sp?.tabColor?.rgb).toBe('FF00B050');
    expect(sp?.outlinePr?.summaryBelow).toBe(false);
    expect(sp?.outlinePr?.summaryRight).toBe(true);
    expect(sp?.outlinePr?.applyStyles).toBe(true);
    expect(sp?.pageSetUpPr?.fitToPage).toBe(true);
    expect(sp?.pageSetUpPr?.autoPageBreaks).toBe(false);
  });

  it('round-trips a theme-bound tabColor with tint', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    setCell(ws, 1, 1, 1);
    ws.sheetProperties = makeSheetProperties({ tabColor: makeColor({ theme: 4, tint: -0.25 }) });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.sheetProperties?.tabColor?.theme).toBe(4);
    expect(ws2.sheetProperties?.tabColor?.tint).toBeCloseTo(-0.25);
  });

  it('does not emit <sheetPr> when sheetProperties is undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'U');
    setCell(ws, 1, 1, 'plain');
    expect(ws.sheetProperties).toBeUndefined();

    const bytes = await workbookToBytes(wb);
    // Reload via library; should produce no sheetProperties since none was set
    // (and the file should not have a leftover <sheetPr/>).
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.sheetProperties).toBeUndefined();
  });
});
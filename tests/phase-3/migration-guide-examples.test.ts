// Smoke tests for the docs/migrate-from-openpyxl.md examples. The
// migration guide is the first stop for users coming from openpyxl,
// so its examples must compile + run against the actual public API.
// A renamed export or shifted signature now fails CI.

import { describe, expect, it } from 'vitest';

describe('migrate-from-openpyxl — public API smoke', () => {
  it('Loading and saving — all four named symbols are exported from openxml-js/node', async () => {
    const node = await import('../../src/node');
    expect(typeof node.fromFile).toBe('function');
    expect(typeof node.loadWorkbook).toBe('function');
    expect(typeof node.saveWorkbook).toBe('function');
    expect(typeof node.toFile).toBe('function');
  });

  it('Cells — setCell / setCellByCoord / iterWorksheetRows / setFormula / makeErrorValue / makeRichText / makeTextRun / makeDurationValue', async () => {
    const main = await import('../../src/index');
    expect(typeof main.setCell).toBe('function');
    expect(typeof main.setCellByCoord).toBe('function');
    expect(typeof main.iterWorksheetRows).toBe('function');
    expect(typeof main.setFormula).toBe('function');
    expect(typeof main.makeErrorValue).toBe('function');
    expect(typeof main.makeRichText).toBe('function');
    expect(typeof main.makeTextRun).toBe('function');
    expect(typeof main.makeDurationValue).toBe('function');
  });

  it('Styles — setCellFont / setCellFill / setCellNumberFormat with (wb, cell, …) signature', async () => {
    const main = await import('../../src/index');
    expect(typeof main.setCellFont).toBe('function');
    expect(typeof main.setCellFill).toBe('function');
    expect(typeof main.setCellNumberFormat).toBe('function');

    // Exercise the canonical migration-guide flow.
    const wb = main.createWorkbook();
    const ws = main.addWorksheet(wb, 'Style');
    main.setCell(ws, 1, 1, 'styled');
    const cell = ws.rows.get(1)?.get(1);
    if (!cell) throw new Error('cell missing');
    main.setCellFont(wb, cell, main.makeFont({ name: 'Arial', size: 14, bold: true }));
    main.setCellFill(
      wb,
      cell,
      main.makePatternFill({ patternType: 'solid', fgColor: main.makeColor({ rgb: 'FFFFFF00' }) }),
    );
    main.setCellNumberFormat(wb, cell, '#,##0.00');
    expect(cell.styleId).toBeGreaterThan(0);
  });

  it('Worksheets — addWorksheet / sheetNames / getActiveSheet / getSheet / removeSheet / mergeCells / setFreezePanes / getMergedCells', async () => {
    const main = await import('../../src/index');
    expect(typeof main.addWorksheet).toBe('function');
    expect(typeof main.sheetNames).toBe('function');
    expect(typeof main.getActiveSheet).toBe('function');
    expect(typeof main.getSheet).toBe('function');
    expect(typeof main.removeSheet).toBe('function');
    expect(typeof main.mergeCells).toBe('function');
    expect(typeof main.setFreezePanes).toBe('function');
    expect(typeof main.getMergedCells).toBe('function');
  });

  it('Streaming write / read — createWriteOnlyWorkbook / loadWorkbookStream', async () => {
    const streaming = await import('../../src/streaming/index');
    expect(typeof streaming.createWriteOnlyWorkbook).toBe('function');
    expect(typeof streaming.loadWorkbookStream).toBe('function');
  });
});

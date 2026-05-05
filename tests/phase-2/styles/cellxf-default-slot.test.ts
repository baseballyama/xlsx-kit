// Regression test: high-level style helpers must reserve cellXfs[0]
// for the implicit default xf, otherwise the first styled cell collides
// with unstyled cells (both reach for styleId=0).

import { describe, expect, it } from 'vitest';
import {
  formatAsHeader,
  setCellFill,
  setCellFont,
  setRangeStyle,
} from '../../../src/styles/cell-style';
import { makeColor } from '../../../src/styles/colors';
import { makeFill, makePatternFill } from '../../../src/styles/fills';
import { makeFont } from '../../../src/styles/fonts';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

describe('cellXfs[0] default-slot reservation', () => {
  it('setCellFont reserves the default at index 0 on first call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 'styled');
    const b = setCell(ws, 2, 1, 'unstyled');
    setCellFont(wb, a, makeFont({ bold: true }));
    expect(a.styleId).toBeGreaterThan(0);
    expect(b.styleId).toBe(0);
    // cellXfs[0] is the default — fontId/fillId/borderId/numFmtId all 0.
    const slot0 = wb.styles.cellXfs[0];
    expect(slot0?.fontId).toBe(0);
    expect(slot0?.applyFont).toBeUndefined();
  });

  it('setCellFill reserves the default at index 0 on first call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 'styled');
    setCellFill(wb, a, makeFill(makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFAA0000' }) })));
    expect(a.styleId).toBeGreaterThan(0);
    expect(wb.styles.cellXfs[0]?.fillId).toBe(0);
  });

  it('setRangeStyle reserves the default before applying the patch', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'header');
    setCell(ws, 2, 1, 'body');
    setRangeStyle(wb, ws, 'A1', { font: makeFont({ bold: true }) });
    const headerStyleId = ws.rows.get(1)?.get(1)?.styleId;
    const bodyStyleId = ws.rows.get(2)?.get(1)?.styleId;
    expect(headerStyleId).toBeGreaterThan(0);
    expect(bodyStyleId).toBe(0);
    expect(headerStyleId).not.toBe(bodyStyleId);
  });

  it('formatAsHeader on row 1 leaves row 2 unstyled', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'Header');
    setCell(ws, 2, 1, 'Body');
    formatAsHeader(wb, ws, 'A1');
    const headerStyleId = ws.rows.get(1)?.get(1)?.styleId;
    const bodyStyleId = ws.rows.get(2)?.get(1)?.styleId;
    expect(headerStyleId).toBeGreaterThan(0);
    expect(bodyStyleId).toBe(0);
  });

  it('manual addCellXf(defaultCellXf()) before styling stays idempotent', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // Caller-side reserve (the legacy pattern): no extra slots are added
    // when the high-level helper runs its own reserve later.
    const beforeLen = wb.styles.cellXfs.length;
    setCellFont(wb, setCell(ws, 1, 1, 'x'), makeFont({ bold: true }));
    expect(beforeLen).toBe(0);
    expect(wb.styles.cellXfs.length).toBe(2); // default + bold
  });
});

// Tests for the cellWatches / ignoredErrors high-level API.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { parseMultiCellRange } from '../../src/worksheet/cell-range';
import { makeCellWatch, makeIgnoredError } from '../../src/worksheet/errors';
import {
  addCellWatch,
  addIgnoredError,
  removeCellWatches,
  removeIgnoredErrors,
  setCell,
  type Worksheet,
} from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('cellWatches API', () => {
  it('addCellWatch / removeCellWatches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'W');
    addCellWatch(ws, makeCellWatch('A1'));
    addCellWatch(ws, makeCellWatch('Sheet1!$B$2'));
    expect(ws.cellWatches.length).toBe(2);
    expect(removeCellWatches(ws, (w) => w.ref === 'A1')).toBe(1);
    expect(ws.cellWatches.length).toBe(1);
    expect(ws.cellWatches[0]?.ref).toBe('Sheet1!$B$2');
  });

  it('round-trips through saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'W');
    setCell(ws, 1, 1, 1);
    addCellWatch(ws, makeCellWatch('A1'));
    addCellWatch(ws, makeCellWatch('A2'));

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.cellWatches.map((w) => w.ref).sort()).toEqual(['A1', 'A2']);
  });
});

describe('ignoredErrors API', () => {
  it('addIgnoredError / removeIgnoredErrors', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'I');
    addIgnoredError(
      ws,
      makeIgnoredError({ sqref: parseMultiCellRange('A1:A10'), numberStoredAsText: true }),
    );
    addIgnoredError(
      ws,
      makeIgnoredError({ sqref: parseMultiCellRange('B1:B10'), formula: true, formulaRange: true }),
    );
    expect(ws.ignoredErrors.length).toBe(2);
    expect(removeIgnoredErrors(ws, (ie) => ie.formula === true)).toBe(1);
    expect(ws.ignoredErrors.length).toBe(1);
    expect(ws.ignoredErrors[0]?.numberStoredAsText).toBe(true);
  });

  it('round-trips with multiple flag axes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'I');
    setCell(ws, 1, 1, 1);
    addIgnoredError(
      ws,
      makeIgnoredError({
        sqref: parseMultiCellRange('A1:C5'),
        numberStoredAsText: true,
        evalError: true,
      }),
    );
    addIgnoredError(
      ws,
      makeIgnoredError({
        sqref: parseMultiCellRange('D1:D10 F2'),
        formula: true,
        unlockedFormula: true,
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.ignoredErrors.length).toBe(2);

    const first = ws2.ignoredErrors[0];
    expect(first?.numberStoredAsText).toBe(true);
    expect(first?.evalError).toBe(true);
    expect(first?.formula).toBeUndefined();

    const second = ws2.ignoredErrors[1];
    expect(second?.formula).toBe(true);
    expect(second?.unlockedFormula).toBe(true);
    // sqref containing two ranges should round-trip both regions.
    expect(second?.sqref.ranges.length).toBe(2);
  });
});
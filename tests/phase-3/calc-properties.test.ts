// Tests for the typed workbook-level <calcPr> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeCalcProperties } from '../../src/workbook/calc-properties';
import { setCell } from '../../src/worksheet/worksheet';

describe('calcPr round-trip', () => {
  it('preserves calcId / calcMode / fullCalcOnLoad / refMode', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    wb.calcProperties = makeCalcProperties({
      calcId: 162913,
      calcMode: 'manual',
      fullCalcOnLoad: true,
      refMode: 'A1',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const cp = wb2.calcProperties;
    expect(cp?.calcId).toBe(162913);
    expect(cp?.calcMode).toBe('manual');
    expect(cp?.fullCalcOnLoad).toBe(true);
    expect(cp?.refMode).toBe('A1');
  });

  it('round-trips iterative calculation settings', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'I');
    setCell(ws, 1, 1, 1);
    wb.calcProperties = makeCalcProperties({
      iterate: true,
      iterateCount: 50,
      iterateDelta: 0.001,
      fullPrecision: false,
      concurrentCalc: true,
      concurrentManualCount: 4,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const cp = wb2.calcProperties;
    expect(cp?.iterate).toBe(true);
    expect(cp?.iterateCount).toBe(50);
    expect(cp?.iterateDelta).toBeCloseTo(0.001);
    expect(cp?.fullPrecision).toBe(false);
    expect(cp?.concurrentCalc).toBe(true);
    expect(cp?.concurrentManualCount).toBe(4);
  });

  it('drops unknown calcMode / refMode enum values silently', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'X');
    setCell(ws, 1, 1, 1);
    // Force an out-of-spec value via cast to make sure the reader
    // ignores it gracefully on the round-trip (writer emits it; reader
    // refuses to set the typed field).
    wb.calcProperties = makeCalcProperties({
      calcMode: 'gibberish' as never,
      refMode: 'A1',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const cp = wb2.calcProperties;
    expect(cp?.calcMode).toBeUndefined();
    expect(cp?.refMode).toBe('A1');
  });

  it('emits no <calcPr/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.calcProperties).toBeUndefined();
  });
});
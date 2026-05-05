// Tests for workbook calcProperties ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import {
  setCalcMode,
  setCalcOnSave,
  setFullCalcOnLoad,
  setFullPrecision,
  setIterativeCalc,
} from '../../src/workbook/calc-properties';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('calcProperties helpers', () => {
  it('setCalcMode lazily creates calcProperties and sets calcMode', () => {
    const wb = createWorkbook();
    setCalcMode(wb, 'manual');
    expect(wb.calcProperties?.calcMode).toBe('manual');
  });

  it('setIterativeCalc enables iteration with Excel-default fallback for count/delta', () => {
    const wb = createWorkbook();
    setIterativeCalc(wb, true);
    expect(wb.calcProperties?.iterate).toBe(true);
    expect(wb.calcProperties?.iterateCount).toBeUndefined();
    expect(wb.calcProperties?.iterateDelta).toBeUndefined();
  });

  it('setIterativeCalc with count + delta', () => {
    const wb = createWorkbook();
    setIterativeCalc(wb, true, { count: 50, delta: 0.0001 });
    expect(wb.calcProperties?.iterateCount).toBe(50);
    expect(wb.calcProperties?.iterateDelta).toBeCloseTo(0.0001);
  });

  it('setIterativeCalc(false) flips iterate to false', () => {
    const wb = createWorkbook();
    setIterativeCalc(wb, true, { count: 100 });
    setIterativeCalc(wb, false);
    expect(wb.calcProperties?.iterate).toBe(false);
    // count is preserved (caller can pass {count: undefined} to keep it)
    expect(wb.calcProperties?.iterateCount).toBe(100);
  });

  it('setCalcOnSave / setFullCalcOnLoad / setFullPrecision toggle independently', () => {
    const wb = createWorkbook();
    setCalcOnSave(wb, true);
    setFullCalcOnLoad(wb, false);
    setFullPrecision(wb, false);
    expect(wb.calcProperties?.calcOnSave).toBe(true);
    expect(wb.calcProperties?.fullCalcOnLoad).toBe(false);
    expect(wb.calcProperties?.fullPrecision).toBe(false);
  });

  it('all helpers compose into a single calcProperties record', () => {
    const wb = createWorkbook();
    setCalcMode(wb, 'manual');
    setIterativeCalc(wb, true, { count: 200, delta: 0.0005 });
    setCalcOnSave(wb, true);
    setFullCalcOnLoad(wb, true);
    setFullPrecision(wb, true);
    expect(wb.calcProperties).toEqual({
      calcMode: 'manual',
      iterate: true,
      iterateCount: 200,
      iterateDelta: 0.0005,
      calcOnSave: true,
      fullCalcOnLoad: true,
      fullPrecision: true,
    });
  });

  it('round-trips through saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setCalcMode(wb, 'manual');
    setIterativeCalc(wb, true, { count: 75, delta: 0.0002 });
    setCalcOnSave(wb, true);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.calcProperties?.calcMode).toBe('manual');
    expect(wb2.calcProperties?.iterate).toBe(true);
    expect(wb2.calcProperties?.iterateCount).toBe(75);
    expect(wb2.calcProperties?.iterateDelta).toBeCloseTo(0.0002);
    expect(wb2.calcProperties?.calcOnSave).toBe(true);
  });
});

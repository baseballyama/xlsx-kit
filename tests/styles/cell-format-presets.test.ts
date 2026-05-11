// Tests for setCellAsCurrency / setCellAsPercent / setCellAsDate /
// setCellAsNumber Excel format-preset helpers.

import { describe, expect, it } from 'vitest';
import {
  getCellNumberFormat,
  setCellAsCurrency,
  setCellAsDate,
  setCellAsNumber,
  setCellAsPercent,
} from '../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('setCellAsCurrency', () => {
  it('default → "$#,##0.00"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234.5);
    setCellAsCurrency(wb, c);
    expect(getCellNumberFormat(wb, c)).toBe('$#,##0.00');
  });

  it('custom symbol + decimals=0', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234);
    setCellAsCurrency(wb, c, { symbol: '¥', decimals: 0 });
    expect(getCellNumberFormat(wb, c)).toBe('¥#,##0');
  });

  it('accounting layout uses Excel "_-$* …" template', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234.5);
    setCellAsCurrency(wb, c, { accounting: true });
    expect(getCellNumberFormat(wb, c)).toBe(
      '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-',
    );
  });
});

describe('setCellAsPercent', () => {
  it('default → "0%"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 0.5);
    setCellAsPercent(wb, c);
    expect(getCellNumberFormat(wb, c)).toBe('0%');
  });

  it('decimals=2 → "0.00%"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 0.5);
    setCellAsPercent(wb, c, 2);
    expect(getCellNumberFormat(wb, c)).toBe('0.00%');
  });

  it('rejects negative / non-integer decimals', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 0.5);
    expect(() => setCellAsPercent(wb, c, -1)).toThrow(/non-negative integer/);
    expect(() => setCellAsPercent(wb, c, 1.5)).toThrow(/non-negative integer/);
  });
});

describe('setCellAsDate', () => {
  it('default → "yyyy-mm-dd"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, new Date('2024-01-02'));
    setCellAsDate(wb, c);
    expect(getCellNumberFormat(wb, c)).toBe('yyyy-mm-dd');
  });

  it('custom format passes through', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, new Date('2024-01-02'));
    setCellAsDate(wb, c, 'm/d/yyyy hh:mm');
    expect(getCellNumberFormat(wb, c)).toBe('m/d/yyyy hh:mm');
  });
});

describe('setCellAsNumber', () => {
  it('default → "#,##0"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234567);
    setCellAsNumber(wb, c);
    expect(getCellNumberFormat(wb, c)).toBe('#,##0');
  });

  it('decimals=2 → "#,##0.00"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234.5);
    setCellAsNumber(wb, c, 2);
    expect(getCellNumberFormat(wb, c)).toBe('#,##0.00');
  });

  it('rejects negative / non-integer decimals', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1234);
    expect(() => setCellAsNumber(wb, c, -2)).toThrow(/non-negative integer/);
    expect(() => setCellAsNumber(wb, c, 0.5)).toThrow(/non-negative integer/);
  });
});

describe('format-preset dedup', () => {
  it('two cells using the same preset share a numFmtId via the xf pool', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 1);
    const b = setCell(ws, 2, 1, 2);
    setCellAsPercent(wb, a, 2);
    setCellAsPercent(wb, b, 2);
    const xfs = wb.styles.cellXfs;
    expect(xfs[a.styleId]?.numFmtId).toBe(xfs[b.styleId]?.numFmtId);
    // The xf pool deduped the format too — both cells share the same xf id.
    expect(a.styleId).toBe(b.styleId);
  });

  it('custom (non-built-in) currency code dedups in the custom numFmts map', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 1);
    const b = setCell(ws, 2, 1, 2);
    setCellAsCurrency(wb, a, { symbol: '€', decimals: 2 });
    setCellAsCurrency(wb, b, { symbol: '€', decimals: 2 });
    const customs = [...wb.styles.numFmts.values()].filter((code) => code === '€#,##0.00');
    expect(customs.length).toBe(1);
  });
});

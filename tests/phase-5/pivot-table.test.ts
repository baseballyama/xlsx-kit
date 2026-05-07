// Tests for pivotTable — row × column aggregation.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { pivotTable, setCell } from '../../src/worksheet/worksheet';

const fillTable = (ws: ReturnType<typeof addWorksheet>) => {
  // | region | category | sales |
  // | east   | A        |  100  |
  // | east   | B        |   50  |
  // | east   | A        |  200  |
  // | west   | A        |   30  |
  // | west   | B        |   70  |
  setCell(ws, 1, 1, 'region');
  setCell(ws, 1, 2, 'category');
  setCell(ws, 1, 3, 'sales');
  setCell(ws, 2, 1, 'east'); setCell(ws, 2, 2, 'A'); setCell(ws, 2, 3, 100);
  setCell(ws, 3, 1, 'east'); setCell(ws, 3, 2, 'B'); setCell(ws, 3, 3, 50);
  setCell(ws, 4, 1, 'east'); setCell(ws, 4, 2, 'A'); setCell(ws, 4, 3, 200);
  setCell(ws, 5, 1, 'west'); setCell(ws, 5, 2, 'A'); setCell(ws, 5, 3, 30);
  setCell(ws, 6, 1, 'west'); setCell(ws, 6, 2, 'B'); setCell(ws, 6, 3, 70);
};

describe('pivotTable', () => {
  it('default aggregate=sum produces row × column totals', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    fillTable(ws);
    expect(
      pivotTable(ws, 'A1:C6', { rowKey: 'region', colKey: 'category', valueKey: 'sales' }),
    ).toEqual({
      east: { A: 300, B: 50 },
      west: { A: 30, B: 70 },
    });
  });

  it('aggregate=count tallies non-null cells (regardless of type)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    fillTable(ws);
    expect(
      pivotTable(ws, 'A1:C6', {
        rowKey: 'region',
        colKey: 'category',
        valueKey: 'sales',
        aggregate: 'count',
      }),
    ).toEqual({
      east: { A: 2, B: 1 },
      west: { A: 1, B: 1 },
    });
  });

  it('aggregate=max returns the largest numeric value per cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    fillTable(ws);
    expect(
      pivotTable(ws, 'A1:C6', {
        rowKey: 'region',
        colKey: 'category',
        valueKey: 'sales',
        aggregate: 'max',
      }),
    ).toEqual({
      east: { A: 200, B: 50 },
      west: { A: 30, B: 70 },
    });
  });

  it('aggregate=mean returns sum/count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    fillTable(ws);
    expect(
      pivotTable(ws, 'A1:C6', {
        rowKey: 'region',
        colKey: 'category',
        valueKey: 'sales',
        aggregate: 'mean',
      }),
    ).toEqual({
      east: { A: 150, B: 50 },
      west: { A: 30, B: 70 },
    });
  });

  it('throws when any of rowKey / colKey / valueKey is missing from the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    fillTable(ws);
    expect(() =>
      pivotTable(ws, 'A1:C6', { rowKey: 'missing', colKey: 'category', valueKey: 'sales' }),
    ).toThrow(/missing/);
    expect(() =>
      pivotTable(ws, 'A1:C6', { rowKey: 'region', colKey: 'missing', valueKey: 'sales' }),
    ).toThrow(/missing/);
    expect(() =>
      pivotTable(ws, 'A1:C6', { rowKey: 'region', colKey: 'category', valueKey: 'missing' }),
    ).toThrow(/missing/);
  });
});

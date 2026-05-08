// Tests for getPopulatedRowIndices / getPopulatedColumnIndices.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getPopulatedColumnIndices,
  getPopulatedRowIndices,
  setCell,
} from '../../src/worksheet/worksheet';

describe('getPopulatedRowIndices', () => {
  it('returns sorted indices of rows with at least one populated cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 1, 'r5');
    setCell(ws, 1, 1, 'r1');
    setCell(ws, 3, 2, 'r3');
    expect(getPopulatedRowIndices(ws)).toEqual([1, 3, 5]);
  });

  it('returns [] for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getPopulatedRowIndices(ws)).toEqual([]);
  });
});

describe('getPopulatedColumnIndices', () => {
  it('returns sorted distinct column indices across all rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 5, 'a');
    setCell(ws, 2, 1, 'b');
    setCell(ws, 3, 3, 'c');
    setCell(ws, 1, 1, 'd');
    expect(getPopulatedColumnIndices(ws)).toEqual([1, 3, 5]);
  });

  it('dedupes columns shared across rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 2, 'a');
    setCell(ws, 5, 2, 'b');
    setCell(ws, 9, 2, 'c');
    expect(getPopulatedColumnIndices(ws)).toEqual([2]);
  });

  it('returns [] for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getPopulatedColumnIndices(ws)).toEqual([]);
  });
});

// Tests for renameSheet / moveSheet helpers.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  moveSheet,
  renameSheet,
  setActiveSheet,
  sheetNames,
} from '../../src/workbook/workbook';

describe('renameSheet', () => {
  it('renames an existing sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Old');
    addWorksheet(wb, 'B');
    renameSheet(wb, 'Old', 'New');
    expect(sheetNames(wb)).toEqual(['New', 'B']);
  });

  it('throws on unknown source', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => renameSheet(wb, 'Missing', 'Other')).toThrow();
  });

  it('throws on duplicate target', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    expect(() => renameSheet(wb, 'A', 'B')).toThrow();
  });

  it('rejects empty target', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => renameSheet(wb, 'A', '')).toThrow();
  });

  it('no-op when oldTitle === newTitle', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    renameSheet(wb, 'A', 'A');
    expect(sheetNames(wb)).toEqual(['A']);
  });

  it('rejects duplicate target case-insensitively', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Alpha');
    addWorksheet(wb, 'Beta');
    expect(() => renameSheet(wb, 'Alpha', 'BETA')).toThrow(/already in use/);
    expect(() => renameSheet(wb, 'Alpha', 'beta')).toThrow(/already in use/);
  });

  it('allows a case-only rename of the same sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    addWorksheet(wb, 'Other');
    renameSheet(wb, 'Data', 'DATA');
    expect(sheetNames(wb)).toEqual(['DATA', 'Other']);
  });
});

describe('moveSheet', () => {
  it('moves a sheet to a new index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    moveSheet(wb, 'A', 2);
    expect(sheetNames(wb)).toEqual(['B', 'C', 'A']);
  });

  it('clamps toIndex to valid range', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    moveSheet(wb, 'A', 99); // clamps to 2
    expect(sheetNames(wb)).toEqual(['B', 'C', 'A']);
    moveSheet(wb, 'A', -5); // clamps to 0
    expect(sheetNames(wb)).toEqual(['A', 'B', 'C']);
  });

  it('keeps activeSheetIndex pointing at the same logical sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setActiveSheet(wb, 'B');
    expect(wb.activeSheetIndex).toBe(1);
    moveSheet(wb, 'B', 0);
    expect(wb.activeSheetIndex).toBe(0);
    expect(sheetNames(wb)).toEqual(['B', 'A', 'C']);
  });

  it('throws on unknown sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => moveSheet(wb, 'Missing', 0)).toThrow();
  });

  it('rejects non-integer index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => moveSheet(wb, 'A', 1.5)).toThrow();
  });
});
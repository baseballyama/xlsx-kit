// Tests for the sheet-protection ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { isSheetProtected, protectSheet, unprotectSheet } from '../../src/xlsx/worksheet/protection';

describe('sheet-protection helpers', () => {
  it('protectSheet applies Excel "Protect Sheet" defaults', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(isSheetProtected(ws)).toBe(false);

    protectSheet(ws);
    expect(isSheetProtected(ws)).toBe(true);
    expect(ws.sheetProtection?.sheet).toBe(true);
    expect(ws.sheetProtection?.objects).toBe(true);
    expect(ws.sheetProtection?.scenarios).toBe(true);
    expect(ws.sheetProtection?.formatCells).toBe(false);
    expect(ws.sheetProtection?.sort).toBe(false);
  });

  it('protectSheet overrides allow specific actions', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    protectSheet(ws, { sort: true, autoFilter: true, selectUnlockedCells: true });
    expect(ws.sheetProtection?.sort).toBe(true);
    expect(ws.sheetProtection?.autoFilter).toBe(true);
    expect(ws.sheetProtection?.selectUnlockedCells).toBe(true);
    // Defaults still applied for the rest.
    expect(ws.sheetProtection?.formatRows).toBe(false);
  });

  it('protectSheet accepts the password-hash quad', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    protectSheet(ws, {
      algorithmName: 'SHA-512',
      hashValue: 'aGFzaA==',
      saltValue: 'c2FsdA==',
      spinCount: 100000,
    });
    expect(ws.sheetProtection?.algorithmName).toBe('SHA-512');
    expect(ws.sheetProtection?.spinCount).toBe(100000);
    expect(ws.sheetProtection?.sheet).toBe(true);
  });

  it('unprotectSheet drops the protection record', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    protectSheet(ws);
    expect(isSheetProtected(ws)).toBe(true);
    unprotectSheet(ws);
    expect(isSheetProtected(ws)).toBe(false);
    expect(ws.sheetProtection).toBeUndefined();
  });
});
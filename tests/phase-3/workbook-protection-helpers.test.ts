// Tests for the workbook-protection ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  isWorkbookProtected,
  protectWorkbook,
  unprotectWorkbook,
} from '../../src/workbook/protection';

describe('workbook-protection helpers', () => {
  it('protectWorkbook sets lockStructure=true by default', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(isWorkbookProtected(wb)).toBe(false);

    protectWorkbook(wb);
    expect(isWorkbookProtected(wb)).toBe(true);
    expect(wb.workbookProtection?.lockStructure).toBe(true);
    expect(wb.workbookProtection?.lockWindows).toBeUndefined();
    expect(wb.workbookProtection?.lockRevision).toBeUndefined();
  });

  it('protectWorkbook overrides lock additional axes', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    protectWorkbook(wb, { lockWindows: true, lockRevision: true });
    expect(wb.workbookProtection?.lockStructure).toBe(true);
    expect(wb.workbookProtection?.lockWindows).toBe(true);
    expect(wb.workbookProtection?.lockRevision).toBe(true);
  });

  it('protectWorkbook accepts password-hash quads', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    protectWorkbook(wb, {
      workbookAlgorithmName: 'SHA-512',
      workbookHashValue: 'aGFzaA==',
      workbookSaltValue: 'c2FsdA==',
      workbookSpinCount: 100000,
    });
    expect(wb.workbookProtection?.workbookAlgorithmName).toBe('SHA-512');
    expect(wb.workbookProtection?.workbookSpinCount).toBe(100000);
    expect(wb.workbookProtection?.lockStructure).toBe(true);
  });

  it('unprotectWorkbook drops the record', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    protectWorkbook(wb);
    expect(isWorkbookProtected(wb)).toBe(true);
    unprotectWorkbook(wb);
    expect(isWorkbookProtected(wb)).toBe(false);
    expect(wb.workbookProtection).toBeUndefined();
  });
});
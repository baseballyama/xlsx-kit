// Tests for workbookProperties ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  setDate1904,
  setFilterPrivacy,
  setUpdateLinksMode,
  setWorkbookCodeName,
} from '../../src/workbook/workbook-properties';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('setWorkbookCodeName', () => {
  it('lazily creates workbookProperties and sets codeName', () => {
    const wb = createWorkbook();
    setWorkbookCodeName(wb, 'ThisWorkbook');
    expect(wb.workbookProperties?.codeName).toBe('ThisWorkbook');
  });

  it('localised non-ASCII codeName round-trips', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setWorkbookCodeName(wb, 'ЭтаКнига');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.workbookProperties?.codeName).toBe('ЭтаКнига');
  });

  it('empty string is preserved as-is', () => {
    const wb = createWorkbook();
    setWorkbookCodeName(wb, '');
    expect(wb.workbookProperties?.codeName).toBe('');
  });
});

describe('setDate1904', () => {
  it('flips both wb.date1904 and the workbookPr mirror', () => {
    const wb = createWorkbook();
    expect(wb.date1904).toBe(false);
    setDate1904(wb, true);
    expect(wb.date1904).toBe(true);
    expect(wb.workbookProperties?.date1904).toBe(true);
  });

  it('round-trip preserves the toggle', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setDate1904(wb, true);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.date1904).toBe(true);
  });
});

describe('setUpdateLinksMode', () => {
  it('accepts userSet / never / always', () => {
    const wb = createWorkbook();
    setUpdateLinksMode(wb, 'never');
    expect(wb.workbookProperties?.updateLinks).toBe('never');
    setUpdateLinksMode(wb, 'always');
    expect(wb.workbookProperties?.updateLinks).toBe('always');
  });
});

describe('setFilterPrivacy', () => {
  it('toggles the filterPrivacy hint', () => {
    const wb = createWorkbook();
    setFilterPrivacy(wb, true);
    expect(wb.workbookProperties?.filterPrivacy).toBe(true);
  });
});

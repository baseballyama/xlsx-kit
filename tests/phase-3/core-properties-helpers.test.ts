// Tests for Office core-property ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  setWorkbookCategory,
  setWorkbookCreator,
  setWorkbookDescription,
  setWorkbookKeywords,
  setWorkbookLastModifiedBy,
  setWorkbookSubject,
  setWorkbookTitle,
} from '../../src/packaging/core';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('coreProperties ergonomic helpers', () => {
  it('lazily allocate wb.properties and write each field', () => {
    const wb = createWorkbook();
    expect(wb.properties).toBeUndefined();
    setWorkbookCreator(wb, 'Alice');
    setWorkbookTitle(wb, 'Quarterly Report');
    setWorkbookSubject(wb, 'Finance');
    setWorkbookDescription(wb, 'Q3 2025 results');
    setWorkbookKeywords(wb, 'finance;quarterly;2025');
    setWorkbookLastModifiedBy(wb, 'Bob');
    setWorkbookCategory(wb, 'Reports');
    expect(wb.properties).toEqual({
      creator: 'Alice',
      title: 'Quarterly Report',
      subject: 'Finance',
      description: 'Q3 2025 results',
      keywords: 'finance;quarterly;2025',
      lastModifiedBy: 'Bob',
      category: 'Reports',
    });
  });

  it('subsequent calls overwrite the prior value', () => {
    const wb = createWorkbook();
    setWorkbookCreator(wb, 'Alice');
    setWorkbookCreator(wb, 'Bob');
    expect(wb.properties?.creator).toBe('Bob');
  });

  it('round-trips through saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setWorkbookCreator(wb, 'Alice');
    setWorkbookTitle(wb, 'Title');
    setWorkbookSubject(wb, 'Subj');
    setWorkbookDescription(wb, 'Desc');
    setWorkbookKeywords(wb, 'k1;k2');
    setWorkbookLastModifiedBy(wb, 'Bob');
    setWorkbookCategory(wb, 'Cat');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.properties?.creator).toBe('Alice');
    expect(wb2.properties?.title).toBe('Title');
    expect(wb2.properties?.subject).toBe('Subj');
    expect(wb2.properties?.description).toBe('Desc');
    expect(wb2.properties?.keywords).toBe('k1;k2');
    expect(wb2.properties?.lastModifiedBy).toBe('Bob');
    expect(wb2.properties?.category).toBe('Cat');
  });

  it('non-ASCII metadata round-trips', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setWorkbookCreator(wb, '山田太郎');
    setWorkbookTitle(wb, '四半期報告 📊');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.properties?.creator).toBe('山田太郎');
    expect(wb2.properties?.title).toBe('四半期報告 📊');
  });
});

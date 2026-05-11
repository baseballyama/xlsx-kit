// Save-time sheet-name guardrails. The public mutators (`addWorksheet`,
// `renameSheet`, ...) already validate. This file covers the failure mode
// where someone bypassed them by mutating `sheet.title` directly: without the
// save-time gate, those bad names would silently produce an xlsx Excel
// refuses to open.

import { describe, expect, it } from 'vitest';
import { workbookToBytes } from '../../src/io/save';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('saveWorkbook — sheet title gate', () => {
  it('rejects a title that exceeds 31 characters', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    // Bypass the public mutator to simulate stale or hand-built state.
    ws.title = 'a'.repeat(32);
    await expect(workbookToBytes(wb)).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });

  it('rejects forbidden Excel characters in titles (: \\ / ? * [ ])', async () => {
    for (const bad of ['has:colon', 'has\\back', 'has/slash', 'has?q', 'has*star', 'has[lb', 'has]rb']) {
      const wb = createWorkbook();
      const ws = addWorksheet(wb, 'Sheet1');
      ws.title = bad;
      await expect(workbookToBytes(wb)).rejects.toThrowError(/must not contain/);
    }
  });

  it('rejects titles that start or end with an apostrophe', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    ws.title = "'leading";
    await expect(workbookToBytes(wb)).rejects.toThrowError(/apostrophe/);
  });

  it('rejects the reserved "History" name (case-insensitive)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    ws.title = 'history';
    await expect(workbookToBytes(wb)).rejects.toThrowError(/History/);
  });

  it('rejects empty titles', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    ws.title = '';
    await expect(workbookToBytes(wb)).rejects.toThrowError(/1\.\.31/);
  });

  it('rejects case-insensitive duplicates introduced via direct mutation', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'Data');
    addWorksheet(wb, 'Other');
    // The public renameSheet would catch this; direct mutation doesn't.
    a.title = 'OTHER';
    await expect(workbookToBytes(wb)).rejects.toThrowError(/collides/);
  });

  it('accepts a workbook with valid, unique titles', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    addWorksheet(wb, 'Summary');
    addWorksheet(wb, 'Data 2025');
    const bytes = await workbookToBytes(wb);
    expect(bytes.byteLength).toBeGreaterThan(0);
  });
});

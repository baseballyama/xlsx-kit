// Tests for the typed Workbook.pivotCaches array.

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

const here = dirname(fileURLToPath(import.meta.url));
const PIVOT_FIXTURE = resolve(here, '../../reference/openpyxl/openpyxl/reader/tests/data/pivot.xlsx');

describe('pivotCaches round-trip', () => {
  it('extracts pivotCaches from the genuine pivot fixture', async () => {
    const original = readFileSync(PIVOT_FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.pivotCaches).toBeDefined();
    expect((wb.pivotCaches ?? []).length).toBeGreaterThan(0);
    for (const pc of wb.pivotCaches ?? []) {
      expect(typeof pc.cacheId).toBe('number');
      expect(pc.rId).toMatch(/^rId/);
    }
  });

  it('survives a load → save → load round-trip with the same cache IDs', async () => {
    const original = readFileSync(PIVOT_FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const idsBefore = (wb.pivotCaches ?? []).map((p) => p.cacheId);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const idsAfter = (wb2.pivotCaches ?? []).map((p) => p.cacheId);

    expect(idsAfter).toEqual(idsBefore);
  });

  it('emits no <pivotCaches/> when undefined on a freshly built workbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.pivotCaches).toBeUndefined();
  });
});
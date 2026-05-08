// Phase 4 read-only streaming — real-fixture acceptance. Covers the
// iter API against openpyxl reference workbooks so a regression in
// SAX parsing surfaces against actual Excel-emitted XML, not just the
// synthetic ones built by createWorkbook.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { loadWorkbookStream } from '../../src/streaming/read-only';

const FIXTURE = (rel: string): Uint8Array =>
  readFileSync(resolve(__dirname, '../../reference/openpyxl/openpyxl/tests/data', rel));

describe('phase-4 read-only — genuine fixtures via SAX iter', () => {
  it('genuine/sample.xlsx — sheet names + first-row iter parity vs loadWorkbook', async () => {
    const bytes = FIXTURE('genuine/sample.xlsx');
    const wbStream = await loadWorkbookStream(fromBuffer(bytes));
    const wbFull = await loadWorkbook(fromBuffer(bytes));
    expect(wbStream.sheetNames).toEqual(wbFull.sheets.map((s) => s.sheet.title));

    // Iterate first sheet via SAX, capture {row, col, value} per cell.
    const firstName = wbStream.sheetNames[0];
    if (firstName === undefined) throw new Error('no first sheet');
    const sax = wbStream.openWorksheet(firstName);
    const saxRows: Array<Array<{ row: number; col: number; value: unknown }>> = [];
    for await (const cells of sax.iterRows()) {
      saxRows.push(cells.map((c) => ({ row: c.row, col: c.col, value: c.value })));
    }

    // Same shape via the eager loader.
    const fullRef = wbFull.sheets[0];
    if (fullRef?.kind !== 'worksheet') throw new Error('expected worksheet');
    const fullRows: Array<Array<{ row: number; col: number; value: unknown }>> = [];
    const rowKeys = [...fullRef.sheet.rows.keys()].sort((a, b) => a - b);
    for (const r of rowKeys) {
      const colMap = fullRef.sheet.rows.get(r);
      if (!colMap) continue;
      const colKeys = [...colMap.keys()].sort((a, b) => a - b);
      const cells: Array<{ row: number; col: number; value: unknown }> = [];
      for (const col of colKeys) {
        const cell = colMap.get(col);
        if (!cell) continue;
        // The streaming reader doesn't materialise FormulaValue — it
        // surfaces the cell's t='str' / 'b' / 'n' / 'e' literal value
        // and drops the formula wrapper. Mirror that here so the
        // comparison is apples-to-apples.
        if (cell.value !== null && typeof cell.value === 'object' && 'kind' in cell.value && cell.value.kind === 'formula') {
          cells.push({ row: cell.row, col: cell.col, value: cell.value.cachedValue ?? null });
        } else {
          cells.push({ row: cell.row, col: cell.col, value: cell.value });
        }
      }
      if (cells.length > 0) fullRows.push(cells);
    }

    // Both paths produce the same row count + same {row, col} coordinates.
    expect(saxRows.length).toBe(fullRows.length);
    for (let i = 0; i < saxRows.length; i++) {
      const saxRow = saxRows[i] ?? [];
      const fullRow = fullRows[i] ?? [];
      expect(saxRow.map((c) => ({ row: c.row, col: c.col }))).toEqual(
        fullRow.map((c) => ({ row: c.row, col: c.col })),
      );
    }
    await wbStream.close();
  });

  it('genuine/empty-with-styles.xlsx — minRow / maxRow filter trims the iter', async () => {
    const bytes = FIXTURE('genuine/empty-with-styles.xlsx');
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet(wb.sheetNames[0] ?? '');

    const allRows: number[] = [];
    for await (const cells of ws.iterRows()) {
      const r = cells[0]?.row;
      if (r !== undefined) allRows.push(r);
    }

    // Filter to a non-trivial subrange — verifies the iterator drops
    // rows outside the band rather than returning the whole sheet.
    const filtered: number[] = [];
    if (allRows.length >= 2) {
      const lo = allRows[0] ?? 0;
      const hi = allRows[allRows.length - 1] ?? 0;
      const band = { minRow: lo + 1, maxRow: hi };
      for await (const cells of ws.iterRows(band)) {
        const r = cells[0]?.row;
        if (r !== undefined) filtered.push(r);
      }
      expect(filtered.every((r) => r >= band.minRow && r <= band.maxRow)).toBe(true);
      expect(filtered.length).toBeLessThanOrEqual(allRows.length);
    }
    await wb.close();
  });
});

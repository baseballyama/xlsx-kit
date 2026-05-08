// Phase 4 §2.x — row-offset index correctness for sub-sheet
// streaming reads. iterRows({minRow}) builds a `<row r="N">` byte
// offset map on first use and binary-searches it instead of
// SAX-walking from start; this file pins the visible behaviour so a
// regression in either the index builder or the saxes wrap doesn't
// silently corrupt the iter output.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { workbookToBytes } from '../../src/io/save';
import { loadWorkbookStream } from '../../src/streaming/read-only';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const buildSheet = async (rows: number): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'A');
  for (let r = 1; r <= rows; r++) {
    setCell(ws, r, 1, r);
    setCell(ws, r, 2, `row-${r}`);
  }
  return workbookToBytes(wb);
};

describe('phase-4 — row-offset index for sub-sheet iter', () => {
  it('iterRows({ minRow: K }) yields only rows ≥ K', async () => {
    const bytes = await buildSheet(50);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 40 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]);
    await wb.close();
  });

  it('iterRows({ minRow, maxRow }) yields the band only', async () => {
    const bytes = await buildSheet(20);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 5, maxRow: 8 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([5, 6, 7, 8]);
    await wb.close();
  });

  it('iterRows({ minRow > maxKnownRow }) yields no rows without throwing', async () => {
    const bytes = await buildSheet(10);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const seen: number[] = [];
    for await (const row of ws.iterRows({ minRow: 999 })) {
      const r = row[0]?.row;
      if (r !== undefined) seen.push(r);
    }
    expect(seen).toEqual([]);
    await wb.close();
  });

  it('the index is reused — second band query on the same worksheet still works', async () => {
    const bytes = await buildSheet(30);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const a: number[] = [];
    for await (const row of ws.iterRows({ minRow: 10, maxRow: 12 })) {
      const r = row[0]?.row;
      if (r !== undefined) a.push(r);
    }
    const b: number[] = [];
    for await (const row of ws.iterRows({ minRow: 25, maxRow: 27 })) {
      const r = row[0]?.row;
      if (r !== undefined) b.push(r);
    }
    expect(a).toEqual([10, 11, 12]);
    expect(b).toEqual([25, 26, 27]);
    await wb.close();
  });

  it('cell values stay correct in the sliced path (sharedStrings + numeric mix)', async () => {
    const bytes = await buildSheet(10);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('A');

    const last: ReadonlyArray<{ col: number; value: unknown }> = await (async () => {
      for await (const row of ws.iterRows({ minRow: 7, maxRow: 7 })) {
        return row.map((c) => ({ col: c.col, value: c.value }));
      }
      return [];
    })();
    expect(last).toEqual([
      { col: 1, value: 7 },
      { col: 2, value: 'row-7' },
    ]);
    await wb.close();
  });
});

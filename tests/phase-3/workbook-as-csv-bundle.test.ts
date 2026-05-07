// Tests for getWorkbookAsCsvBundle — zip of <title>.csv entries.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsCsvBundle,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const decode = (entries: Record<string, Uint8Array>): Record<string, string> => {
  const out: Record<string, string> = {};
  const dec = new TextDecoder();
  for (const [k, v] of Object.entries(entries)) out[k] = dec.decode(v);
  return out;
};

describe('getWorkbookAsCsvBundle', () => {
  it('returns a valid zip with one <title>.csv entry per worksheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(b, 1, 1, 'b1');
    const bytes = getWorkbookAsCsvBundle(wb);
    const entries = decode(unzipSync(bytes));
    expect(Object.keys(entries).sort()).toEqual(['A.csv', 'B.csv']);
    expect(entries['A.csv']).toBe('a1');
    expect(entries['B.csv']).toBe('b1');
  });

  it('sanitises titles with filesystem-unfriendly characters', () => {
    const wb = createWorkbook();
    // Excel disallows : \ / ? * [ ] but permits | < > , ; etc.
    addWorksheet(wb, 'a|b<c>');
    const entries = decode(unzipSync(getWorkbookAsCsvBundle(wb)));
    expect(Object.keys(entries)).toEqual(['a_b_c_.csv']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    addChartsheet(wb, 'Chart1');
    const entries = decode(unzipSync(getWorkbookAsCsvBundle(wb)));
    expect(Object.keys(entries)).toEqual(['Data.csv']);
  });

  it('forwards opts.delimiter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    const entries = decode(unzipSync(getWorkbookAsCsvBundle(wb, { delimiter: ';' })));
    expect(entries['S.csv']).toBe('a;b');
  });

  it('returns a valid (empty) zip for a workbook with no worksheets', () => {
    const wb = createWorkbook();
    const bytes = getWorkbookAsCsvBundle(wb);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(0); // empty zip still has end-of-central-directory
    expect(Object.keys(unzipSync(bytes))).toEqual([]);
  });
});

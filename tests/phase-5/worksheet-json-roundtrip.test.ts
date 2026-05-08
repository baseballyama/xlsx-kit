// End-to-end round-trip: worksheetToJson → parseJsonToRange.
// Catches drift between the export coercion and the import coercion.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { parseJsonToRange, worksheetToJson } from '../../src/worksheet/json';
import { getCell, setCell } from '../../src/worksheet/worksheet';

describe('worksheetToJson + parseJsonToRange round-trip', () => {
  it('round-trips a basic header + data block', () => {
    const wb = createWorkbook();
    const src = addWorksheet(wb, 'src');
    setCell(src, 1, 1, 'name'); setCell(src, 1, 2, 'age');
    setCell(src, 2, 1, 'Alice'); setCell(src, 2, 2, 30);
    setCell(src, 3, 1, 'Bob'); setCell(src, 3, 2, 25);
    const json = worksheetToJson(src, 'A1:B3');

    const dst = addWorksheet(wb, 'dst');
    parseJsonToRange(dst, 'A1', json);

    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expect(getCell(dst, r, c)?.value).toEqual(getCell(src, r, c)?.value);
      }
    }
  });

  it('round-trips Date cells via ISO 8601 string', () => {
    const wb = createWorkbook();
    const src = addWorksheet(wb, 'src');
    const when = new Date('2026-01-15T03:30:00.000Z');
    setCell(src, 1, 1, 'when');
    setCell(src, 2, 1, when);
    const json = worksheetToJson(src, 'A1:A2');

    const dst = addWorksheet(wb, 'dst');
    parseJsonToRange(dst, 'A1', json);
    const restored = getCell(dst, 2, 1)?.value;
    expect(restored).toBeInstanceOf(Date);
    expect((restored as Date).toISOString()).toBe(when.toISOString());
  });
});

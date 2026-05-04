import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import {
  addWorksheet,
  createWorkbook,
  getActiveSheet,
  getSheet,
  getSheetByIndex,
  jsonReplacer,
  jsonReviver,
  removeSheet,
  setActiveSheet,
  sheetNames,
} from '../../src/workbook/workbook';
import {
  appendRow,
  countCells,
  deleteCell,
  getCell,
  getMaxCol,
  getMaxRow,
  iterRows,
  iterValues,
  setCell,
} from '../../src/worksheet/worksheet';

describe('createWorkbook', () => {
  it('starts empty with a fresh Stylesheet', () => {
    const wb = createWorkbook();
    expect(wb.sheets).toEqual([]);
    expect(wb.activeSheetIndex).toBe(0);
    expect(wb.styles.fonts.length).toBe(1);
    expect(wb.styles.fills.length).toBe(2);
    expect(wb.date1904).toBe(false);
  });

  it('honours the date1904 option', () => {
    expect(createWorkbook({ date1904: true }).date1904).toBe(true);
  });
});

describe('addWorksheet', () => {
  it('allocates the next available sheetId starting at 1', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    addWorksheet(wb, 'Sheet2');
    addWorksheet(wb, 'Sheet3');
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([1, 2, 3]);
    expect(sheetNames(wb)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
  });

  it('rejects duplicate titles', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => addWorksheet(wb, 'Data')).toThrowError(OpenXmlSchemaError);
  });

  it('rejects empty / overlong titles', () => {
    const wb = createWorkbook();
    expect(() => addWorksheet(wb, '')).toThrowError(OpenXmlSchemaError);
    expect(() => addWorksheet(wb, 'a'.repeat(32))).toThrowError(OpenXmlSchemaError);
  });

  it('inserts at a specific index when requested', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'C');
    addWorksheet(wb, 'B', { index: 1 });
    expect(sheetNames(wb)).toEqual(['A', 'B', 'C']);
  });

  it('honours the visibility state', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Hidden', { state: 'hidden' });
    expect(wb.sheets[0]?.state).toBe('hidden');
  });

  it('rejects out-of-range insertion indices', () => {
    const wb = createWorkbook();
    expect(() => addWorksheet(wb, 'X', { index: 5 })).toThrowError(OpenXmlSchemaError);
    expect(() => addWorksheet(wb, 'Y', { index: -1 })).toThrowError(OpenXmlSchemaError);
  });
});

describe('getSheet / getSheetByIndex / removeSheet', () => {
  it('looks up by title and by index', () => {
    const wb = createWorkbook();
    const s1 = addWorksheet(wb, 'A');
    const s2 = addWorksheet(wb, 'B');
    expect(getSheet(wb, 'A')).toBe(s1);
    expect(getSheetByIndex(wb, 1)).toBe(s2);
    expect(getSheet(wb, 'Missing')).toBeUndefined();
    expect(getSheetByIndex(wb, 99)).toBeUndefined();
  });

  it('removeSheet drops the matching entry and clamps activeSheetIndex', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setActiveSheet(wb, 'C');
    expect(wb.activeSheetIndex).toBe(2);
    removeSheet(wb, 'C');
    // Deleted the active one — clamps to last remaining (index 1).
    expect(wb.activeSheetIndex).toBe(1);
    expect(sheetNames(wb)).toEqual(['A', 'B']);
  });

  it('removeSheet on a missing title is a no-op', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    removeSheet(wb, 'Z');
    expect(sheetNames(wb)).toEqual(['A']);
  });
});

describe('setActiveSheet / getActiveSheet', () => {
  it('throws on unknown title', () => {
    const wb = createWorkbook();
    expect(() => setActiveSheet(wb, 'Imaginary')).toThrowError(OpenXmlSchemaError);
  });

  it('round-trips through the active index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setActiveSheet(wb, 'B');
    expect(getActiveSheet(wb)).toBe(b);
  });
});

describe('Worksheet getCell / setCell / deleteCell', () => {
  it('setCell populates the sparse map and getCell returns it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const cell = setCell(ws, 1, 1, 42);
    expect(cell.value).toBe(42);
    expect(getCell(ws, 1, 1)).toBe(cell);
  });

  it('updating an existing cell preserves identity', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const a = setCell(ws, 1, 1, 1);
    const b = setCell(ws, 1, 1, 2);
    expect(a).toBe(b);
    expect(b.value).toBe(2);
  });

  it('deleteCell prunes empty rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 5, 3, 'x');
    expect(ws.rows.has(5)).toBe(true);
    deleteCell(ws, 5, 3);
    expect(ws.rows.has(5)).toBe(false);
  });

  it('rejects out-of-range row / col', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(() => setCell(ws, 0, 1, 1)).toThrowError(OpenXmlSchemaError);
    expect(() => setCell(ws, 1, 16385, 1)).toThrowError(OpenXmlSchemaError);
  });
});

describe('appendRow', () => {
  it('places values starting at the next empty row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(appendRow(ws, ['a', 'b', 'c'])).toBe(1);
    expect(appendRow(ws, [1, 2, 3])).toBe(2);
    expect(getCell(ws, 1, 1)?.value).toBe('a');
    expect(getCell(ws, 2, 2)?.value).toBe(2);
  });

  it('skips null / undefined entries', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    appendRow(ws, ['x', null, 'z']);
    expect(getCell(ws, 1, 1)?.value).toBe('x');
    expect(getCell(ws, 1, 2)).toBeUndefined();
    expect(getCell(ws, 1, 3)?.value).toBe('z');
  });

  it('advances the cursor even when all values are empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(appendRow(ws, [])).toBe(1);
    expect(appendRow(ws, ['x'])).toBe(2);
    expect(getCell(ws, 2, 1)?.value).toBe('x');
  });
});

describe('iterRows / iterValues', () => {
  it('iterates rows in ascending order, skipping empty rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 3, 1, 'b'); // gap on row 2
    setCell(ws, 3, 2, 'c');
    const rows = [...iterRows(ws)].map((row) => row.map((c) => c.value));
    expect(rows).toEqual([['a'], ['b', 'c']]);
  });

  it('honours minRow / maxRow / minCol / maxCol filters', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 5, 'b');
    setCell(ws, 2, 3, 'c');
    setCell(ws, 5, 5, 'd');
    const rows = [...iterRows(ws, { minRow: 1, maxRow: 2, minCol: 1, maxCol: 4 })].map((row) =>
      row.map((c) => c.value),
    );
    // row1 col1 only; row2 col3 only; row5 filtered out.
    expect(rows).toEqual([['a'], ['c']]);
  });

  it('iterValues yields the cell values directly', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    appendRow(ws, [1, 2, 3]);
    appendRow(ws, [4, 5, 6]);
    expect([...iterValues(ws)]).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
  });
});

describe('getMaxRow / getMaxCol / countCells', () => {
  it('reflect the actual populated range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 7, 5, 'b');
    setCell(ws, 3, 9, 'c');
    expect(getMaxRow(ws)).toBe(7);
    expect(getMaxCol(ws)).toBe(9);
    expect(countCells(ws)).toBe(3);
  });
});

describe('JSON round-trip via jsonReplacer / jsonReviver', () => {
  it('preserves the rows Map across stringify/parse', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'hello');
    setCell(ws, 2, 1, 42);

    const text = JSON.stringify(wb, jsonReplacer);
    const wb2 = JSON.parse(text, jsonReviver) as typeof wb;

    const restored = wb2.sheets[0]?.sheet;
    expect(restored?.title).toBe('S');
    expect(restored?.rows instanceof Map).toBe(true);
    expect(restored?.rows.get(1) instanceof Map).toBe(true);
    expect(restored?.rows.get(1)?.get(1)?.value).toBe('hello');
    expect(restored?.rows.get(2)?.get(1)?.value).toBe(42);
  });
});

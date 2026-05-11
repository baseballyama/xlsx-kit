import { describe, expect, it } from 'vitest';
import {
  boundariesToRangeString,
  columnIndexFromLetter,
  columnLetterFromIndex,
  coordinateFromString,
  coordinateToTuple,
  MAX_COL,
  MAX_ROW,
  parseSheetRange,
  rangeBoundaries,
  tupleToCoordinate,
} from '../../../src/utils/coordinate';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';

describe('columnLetterFromIndex / columnIndexFromLetter', () => {
  // Verified against openpyxl's get_column_letter / column_index_from_string.
  const cases: ReadonlyArray<readonly [number, string]> = [
    [1, 'A'],
    [2, 'B'],
    [26, 'Z'],
    [27, 'AA'],
    [52, 'AZ'],
    [53, 'BA'],
    [702, 'ZZ'],
    [703, 'AAA'],
    [1000, 'ALL'],
    [16384, 'XFD'],
  ];

  it.each(cases)('columnLetterFromIndex(%i) === %s', (idx, letter) => {
    expect(columnLetterFromIndex(idx)).toBe(letter);
  });

  it.each(cases)('columnIndexFromLetter(%s) === %i', (idx, letter) => {
    expect(columnIndexFromLetter(letter)).toBe(idx);
  });

  it('round-trips every column index', () => {
    for (let i = 1; i <= MAX_COL; i++) {
      expect(columnIndexFromLetter(columnLetterFromIndex(i))).toBe(i);
    }
  });

  it('accepts lower-case and normalises through the cache', () => {
    expect(columnIndexFromLetter('a')).toBe(1);
    expect(columnIndexFromLetter('xfd')).toBe(MAX_COL);
  });

  it('rejects out-of-range indices', () => {
    expect(() => columnLetterFromIndex(0)).toThrowError(OpenXmlSchemaError);
    expect(() => columnLetterFromIndex(MAX_COL + 1)).toThrowError(OpenXmlSchemaError);
    expect(() => columnLetterFromIndex(1.5)).toThrowError(OpenXmlSchemaError);
  });

  it('rejects malformed letter inputs', () => {
    expect(() => columnIndexFromLetter('')).toThrowError(OpenXmlSchemaError);
    expect(() => columnIndexFromLetter('AAAA')).toThrowError(OpenXmlSchemaError);
    expect(() => columnIndexFromLetter('A1')).toThrowError(OpenXmlSchemaError);
    // ZZZ would expand to 18278 > MAX_COL.
    expect(() => columnIndexFromLetter('ZZZ')).toThrowError(OpenXmlSchemaError);
  });
});

describe('coordinateFromString', () => {
  it('splits "A1" into column letter and row', () => {
    expect(coordinateFromString('A1')).toEqual({ column: 'A', row: 1 });
  });

  it('handles absolute markers', () => {
    expect(coordinateFromString('$A$1')).toEqual({ column: 'A', row: 1 });
    expect(coordinateFromString('$AB$42')).toEqual({ column: 'AB', row: 42 });
  });

  it('uppercases the column letter', () => {
    expect(coordinateFromString('xfd1048576')).toEqual({ column: 'XFD', row: 1048576 });
  });

  it('rejects malformed coordinates', () => {
    expect(() => coordinateFromString('A')).toThrowError(OpenXmlSchemaError);
    expect(() => coordinateFromString('1A')).toThrowError(OpenXmlSchemaError);
    expect(() => coordinateFromString('A0')).toThrowError(OpenXmlSchemaError);
    expect(() => coordinateFromString('AAAA1')).toThrowError(OpenXmlSchemaError);
  });
});

describe('coordinateToTuple / tupleToCoordinate', () => {
  it('coordinateToTuple uses 1-based numeric column', () => {
    expect(coordinateToTuple('A1')).toEqual({ col: 1, row: 1 });
    expect(coordinateToTuple('XFD1048576')).toEqual({ col: MAX_COL, row: MAX_ROW });
  });

  it('tupleToCoordinate is the inverse', () => {
    expect(tupleToCoordinate(1, 1)).toBe('A1');
    expect(tupleToCoordinate(27, 100)).toBe('AA100');
    expect(tupleToCoordinate(MAX_COL, MAX_ROW)).toBe('XFD1048576');
  });

  it('tupleToCoordinate rejects bad rows', () => {
    expect(() => tupleToCoordinate(1, 0)).toThrowError(OpenXmlSchemaError);
    expect(() => tupleToCoordinate(1, MAX_ROW + 1)).toThrowError(OpenXmlSchemaError);
  });
});

describe('rangeBoundaries', () => {
  it('parses a rectangular range', () => {
    expect(rangeBoundaries('A1:C3')).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 3 });
    expect(rangeBoundaries('B5:D7')).toEqual({ minCol: 2, minRow: 5, maxCol: 4, maxRow: 7 });
  });

  it('treats single-cell input as a 1×1 range', () => {
    expect(rangeBoundaries('B5')).toEqual({ minCol: 2, minRow: 5, maxCol: 2, maxRow: 5 });
  });

  it('handles whole-column ranges (A:A, A:C)', () => {
    expect(rangeBoundaries('A:A')).toEqual({ minCol: 1, minRow: 1, maxCol: 1, maxRow: MAX_ROW });
    expect(rangeBoundaries('A:C')).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: MAX_ROW });
  });

  it('handles whole-row ranges (1:1, 2:5)', () => {
    expect(rangeBoundaries('1:1')).toEqual({ minCol: 1, minRow: 1, maxCol: MAX_COL, maxRow: 1 });
    expect(rangeBoundaries('2:5')).toEqual({ minCol: 1, minRow: 2, maxCol: MAX_COL, maxRow: 5 });
  });

  it('normalises reversed bounds', () => {
    expect(rangeBoundaries('C3:A1')).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 3 });
    expect(rangeBoundaries('C:A')).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: MAX_ROW });
  });

  it('strips absolute markers transparently', () => {
    expect(rangeBoundaries('$A$1:$B$2')).toEqual({ minCol: 1, minRow: 1, maxCol: 2, maxRow: 2 });
  });

  it('rejects malformed ranges', () => {
    expect(() => rangeBoundaries('')).toThrowError(OpenXmlSchemaError);
    expect(() => rangeBoundaries('A:1')).toThrowError(OpenXmlSchemaError);
    expect(() => rangeBoundaries('Sheet1!A1')).toThrowError(OpenXmlSchemaError);
  });
});

describe('boundariesToRangeString', () => {
  it('emits a single coordinate for 1×1 boundaries', () => {
    expect(boundariesToRangeString({ minCol: 1, minRow: 1, maxCol: 1, maxRow: 1 })).toBe('A1');
  });

  it('emits a range for larger boundaries', () => {
    expect(boundariesToRangeString({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 3 })).toBe('A1:C3');
  });

  it('round-trips with rangeBoundaries', () => {
    for (const r of ['A1', 'A1:B2', 'C5:D6', 'AA1:AB100']) {
      expect(boundariesToRangeString(rangeBoundaries(r))).toBe(r);
    }
  });
});

describe('parseSheetRange', () => {
  it('parses bare sheet names', () => {
    const out = parseSheetRange('Sheet1!A1:B5');
    expect(out.sheet).toBe('Sheet1');
    expect(out.range).toBe('A1:B5');
    expect(out.bounds).toEqual({ minCol: 1, minRow: 1, maxCol: 2, maxRow: 5 });
  });

  it('parses single-quoted sheet names containing spaces', () => {
    const out = parseSheetRange("'Quarter 1'!A1");
    expect(out.sheet).toBe('Quarter 1');
    expect(out.range).toBe('A1');
  });

  it('unescapes doubled single quotes in quoted sheet names', () => {
    const out = parseSheetRange("'Bob''s Data'!B2:C3");
    expect(out.sheet).toBe("Bob's Data");
  });

  it('throws when there is no "!"', () => {
    expect(() => parseSheetRange('Sheet1A1:B5')).toThrowError(OpenXmlSchemaError);
  });
});

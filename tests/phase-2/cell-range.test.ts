import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import {
  intersectionRange,
  iterRangeCoordinates,
  makeCellRange,
  makeMultiCellRange,
  multiCellRangeArea,
  multiCellRangeContainsCell,
  multiCellRangeToString,
  parseMultiCellRange,
  parseRange,
  rangeArea,
  rangeContainsCell,
  rangeContainsRange,
  rangesOverlap,
  rangeToString,
  shiftRange,
  unionRange,
} from '../../src/xlsx/worksheet/cell-range';

describe('makeCellRange / parseRange / rangeToString', () => {
  it('makeCellRange normalises reversed bounds', () => {
    expect(makeCellRange(5, 3, 1, 1)).toEqual({ minRow: 1, minCol: 1, maxRow: 5, maxCol: 3 });
  });

  it('rejects non-integer or out-of-range bounds', () => {
    expect(() => makeCellRange(1.5, 1, 1, 1)).toThrowError(OpenXmlSchemaError);
    expect(() => makeCellRange(1, 1, 1, 99999)).toThrowError(OpenXmlSchemaError);
  });

  it('parseRange / rangeToString round-trip', () => {
    for (const r of ['A1', 'A1:B5', 'AA10:AB20']) {
      expect(rangeToString(parseRange(r))).toBe(r);
    }
  });
});

describe('rangeContainsCell / rangeContainsRange', () => {
  const r = makeCellRange(2, 2, 5, 5);

  it('rangeContainsCell is inclusive on every edge', () => {
    expect(rangeContainsCell(r, 2, 2)).toBe(true);
    expect(rangeContainsCell(r, 5, 5)).toBe(true);
    expect(rangeContainsCell(r, 1, 2)).toBe(false);
    expect(rangeContainsCell(r, 2, 6)).toBe(false);
  });

  it('rangeContainsRange checks both axes', () => {
    expect(rangeContainsRange(r, makeCellRange(3, 3, 4, 4))).toBe(true);
    expect(rangeContainsRange(r, makeCellRange(2, 2, 5, 5))).toBe(true);
    expect(rangeContainsRange(r, makeCellRange(1, 2, 5, 5))).toBe(false);
    expect(rangeContainsRange(r, makeCellRange(3, 3, 6, 4))).toBe(false);
  });
});

describe('shiftRange', () => {
  it('shifts both row and column bounds', () => {
    const r = makeCellRange(1, 1, 2, 2);
    expect(shiftRange(r, 4, 2)).toEqual({ minRow: 5, minCol: 3, maxRow: 6, maxCol: 4 });
    expect(shiftRange(r, -0, -0)).toEqual(r);
  });

  it('rejects non-integer offsets', () => {
    expect(() => shiftRange(makeCellRange(1, 1, 1, 1), 0.5, 0)).toThrowError(OpenXmlSchemaError);
  });
});

describe('union / intersection / overlap / area', () => {
  it('unionRange returns the bounding box', () => {
    const a = makeCellRange(1, 1, 3, 3);
    const b = makeCellRange(5, 5, 7, 7);
    expect(unionRange(a, b)).toEqual({ minRow: 1, minCol: 1, maxRow: 7, maxCol: 7 });
  });

  it('intersectionRange returns null when disjoint', () => {
    const a = makeCellRange(1, 1, 2, 2);
    const b = makeCellRange(3, 3, 4, 4);
    expect(intersectionRange(a, b)).toBeNull();
    expect(rangesOverlap(a, b)).toBe(false);
  });

  it('intersectionRange yields the overlap rectangle', () => {
    const a = makeCellRange(1, 1, 5, 5);
    const b = makeCellRange(3, 3, 7, 7);
    expect(intersectionRange(a, b)).toEqual({ minRow: 3, minCol: 3, maxRow: 5, maxCol: 5 });
    expect(rangesOverlap(a, b)).toBe(true);
  });

  it('rangeArea counts cells inclusive on both axes', () => {
    expect(rangeArea(makeCellRange(1, 1, 1, 1))).toBe(1);
    expect(rangeArea(makeCellRange(1, 1, 2, 3))).toBe(6);
  });
});

describe('iterRangeCoordinates', () => {
  it('yields cells row-major', () => {
    const r = makeCellRange(1, 1, 2, 3);
    expect([...iterRangeCoordinates(r)]).toEqual([
      { row: 1, col: 1 },
      { row: 1, col: 2 },
      { row: 1, col: 3 },
      { row: 2, col: 1 },
      { row: 2, col: 2 },
      { row: 2, col: 3 },
    ]);
  });

  it('count matches rangeArea', () => {
    const r = makeCellRange(5, 5, 10, 10);
    expect([...iterRangeCoordinates(r)].length).toBe(rangeArea(r));
  });
});

describe('MultiCellRange', () => {
  it('parseMultiCellRange splits on whitespace and parses each piece', () => {
    const m = parseMultiCellRange('A1:B2 D5 E10:F20');
    expect(m.ranges).toEqual([parseRange('A1:B2'), parseRange('D5'), parseRange('E10:F20')]);
  });

  it('multiCellRangeToString round-trips through parseMultiCellRange', () => {
    const s = 'A1:B2 D5 E10:F20';
    expect(multiCellRangeToString(parseMultiCellRange(s))).toBe(s);
  });

  it('makeMultiCellRange copies the ranges array', () => {
    const ranges = [makeCellRange(1, 1, 2, 2)];
    const m = makeMultiCellRange(ranges);
    expect(m.ranges).toEqual(ranges);
    expect(m.ranges).not.toBe(ranges);
  });

  it('multiCellRangeContainsCell hits any sub-range', () => {
    const m = parseMultiCellRange('A1:B2 D5');
    expect(multiCellRangeContainsCell(m, 1, 1)).toBe(true);
    expect(multiCellRangeContainsCell(m, 2, 2)).toBe(true);
    expect(multiCellRangeContainsCell(m, 5, 4)).toBe(true);
    expect(multiCellRangeContainsCell(m, 3, 3)).toBe(false);
  });

  it('multiCellRangeArea sums per-range areas (overlaps not deduped)', () => {
    const m = parseMultiCellRange('A1:B2 A1:A2');
    expect(multiCellRangeArea(m)).toBe(4 + 2);
  });
});

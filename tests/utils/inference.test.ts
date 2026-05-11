import { describe, expect, it } from 'vitest';
import { ERROR_CODES, inferCellType } from '../../src/utils/inference';

describe('inferCellType', () => {
  it('booleans → "b"', () => {
    expect(inferCellType(true)).toBe('b');
    expect(inferCellType(false)).toBe('b');
  });

  it('numbers → "n"', () => {
    expect(inferCellType(0)).toBe('n');
    expect(inferCellType(42)).toBe('n');
    expect(inferCellType(3.14)).toBe('n');
    expect(inferCellType(-1.5)).toBe('n');
    // Inferred regardless of whether the cell is later styled as a date.
    expect(inferCellType(44927)).toBe('n');
  });

  it('Date → "d"', () => {
    expect(inferCellType(new Date())).toBe('d');
  });

  it('formula strings → "f"', () => {
    expect(inferCellType('=SUM(A1:A10)')).toBe('f');
    expect(inferCellType('=A1+B1')).toBe('f');
  });

  it('error codes → "e"', () => {
    for (const e of ERROR_CODES) expect(inferCellType(e)).toBe('e');
  });

  it('plain strings → "s"', () => {
    expect(inferCellType('hello')).toBe('s');
    expect(inferCellType('')).toBe('s');
    // Strings that *look* like numbers are still strings.
    expect(inferCellType('42')).toBe('s');
  });

  it('null / undefined → "n" (empty cell)', () => {
    expect(inferCellType(null)).toBe('n');
    expect(inferCellType(undefined)).toBe('n');
  });
});

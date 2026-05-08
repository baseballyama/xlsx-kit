// Tests for the `isMergedCell(c)` type guard.

import { describe, expect, it } from 'vitest';
import type { Cell, MergedCell } from '../../src/xlsx/cell';
import { isMergedCell, makeCell } from '../../src/xlsx/cell';

describe('isMergedCell', () => {
  it('returns false for a regular cell with no `merged` flag', () => {
    const c = makeCell(1, 1, 'hello');
    expect(isMergedCell(c)).toBe(false);
  });

  it('returns true when `merged` is true (the placeholder marker)', () => {
    const c = makeCell(1, 1, 'hello') as MergedCell;
    c.merged = true;
    expect(isMergedCell(c)).toBe(true);
  });

  it('narrows the cell to `MergedCell` so `merged` is statically true after the guard', () => {
    const c: Cell = makeCell(1, 1, 'hello');
    (c as MergedCell).merged = true;
    if (isMergedCell(c)) {
      // `merged` typed `true` (literal) under the narrowed branch.
      const merged: true = c.merged;
      expect(merged).toBe(true);
    } else {
      throw new Error('expected the type guard to narrow into the merged branch');
    }
  });
});

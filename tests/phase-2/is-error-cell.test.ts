// Tests for the `isErrorCell(c)` predicate.

import { describe, expect, it } from 'vitest';
import { isErrorCell, makeCell, makeErrorValue, setFormula } from '../../src/cell';

describe('isErrorCell', () => {
  it('returns true when the cell holds an Excel error value', () => {
    const c = makeCell(1, 1, makeErrorValue('#REF!'));
    expect(isErrorCell(c)).toBe(true);
  });

  it('returns false for a regular string cell', () => {
    const c = makeCell(1, 1, 'hello');
    expect(isErrorCell(c)).toBe(false);
  });

  it('returns false for a numeric cell', () => {
    const c = makeCell(1, 1, 42);
    expect(isErrorCell(c)).toBe(false);
  });

  it('returns false for a formula cell — the value kind is `formula`, not `error`', () => {
    const c = makeCell(1, 1, '');
    setFormula(c, 'A1+1');
    expect(isErrorCell(c)).toBe(false);
  });
});

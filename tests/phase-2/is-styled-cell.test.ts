// Tests for the public `isStyledCell(c)` predicate.

import { describe, expect, it } from 'vitest';
import { isStyledCell, makeCell } from '../../src/cell';

describe('isStyledCell', () => {
  it('returns false when styleId is 0 (the default xf)', () => {
    const c = makeCell(1, 1, 'hello');
    c.styleId = 0;
    expect(isStyledCell(c)).toBe(false);
  });

  it('returns true when styleId is a non-zero index', () => {
    const c = makeCell(1, 1, 'hello');
    c.styleId = 1;
    expect(isStyledCell(c)).toBe(true);
  });

  it('returns false for a freshly-made cell (default styleId is 0)', () => {
    const c = makeCell(1, 1, 'hello');
    expect(isStyledCell(c)).toBe(false);
  });
});

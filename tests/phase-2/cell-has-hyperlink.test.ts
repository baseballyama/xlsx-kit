// Tests for the public `cellHasHyperlink(c)` predicate.

import { describe, expect, it } from 'vitest';
import { cellHasHyperlink, makeCell } from '../../src/cell';

describe('cellHasHyperlink', () => {
  it('returns false when the cell has no hyperlinkId set', () => {
    const c = makeCell(1, 1, 'hello');
    expect(cellHasHyperlink(c)).toBe(false);
  });

  it('returns true when hyperlinkId is set to a positive number', () => {
    const c = makeCell(1, 1, 'hello');
    c.hyperlinkId = 1;
    expect(cellHasHyperlink(c)).toBe(true);
  });

  it('returns true even when hyperlinkId is 0 (a valid hyperlink registry id)', () => {
    const c = makeCell(1, 1, 'hello');
    c.hyperlinkId = 0;
    expect(cellHasHyperlink(c)).toBe(true);
  });
});

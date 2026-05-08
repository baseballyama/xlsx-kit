// Tests for the public `cellHasComment(c)` predicate.

import { describe, expect, it } from 'vitest';
import { cellHasComment, makeCell } from '../../src/cell';

describe('cellHasComment', () => {
  it('returns false when the cell has no commentId set', () => {
    const c = makeCell(1, 1, 'hello');
    expect(cellHasComment(c)).toBe(false);
  });

  it('returns true when commentId is set to a positive number', () => {
    const c = makeCell(1, 1, 'hello');
    c.commentId = 1;
    expect(cellHasComment(c)).toBe(true);
  });

  it('returns true even when commentId is 0 (a valid comment registry id)', () => {
    const c = makeCell(1, 1, 'hello');
    c.commentId = 0;
    expect(cellHasComment(c)).toBe(true);
  });
});

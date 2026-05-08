// Tests for the public `richTextRun` alias of `makeTextRun`.

import { describe, expect, it } from 'vitest';
import { makeTextRun, richTextRun } from '../../src/cell';

describe('richTextRun', () => {
  it('produces a frozen TextRun without font when called with text only', () => {
    const run = richTextRun('hello');
    expect(run.text).toBe('hello');
    expect(run.font).toBeUndefined();
    expect(Object.isFrozen(run)).toBe(true);
    expect(richTextRun).toBe(makeTextRun);
  });

  it('produces a TextRun carrying the supplied font', () => {
    const run = richTextRun('bold', { b: true, sz: 12 });
    expect(run.text).toBe('bold');
    expect(run.font).toEqual({ b: true, sz: 12 });
    expect(Object.isFrozen(run)).toBe(true);
  });
});

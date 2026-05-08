// Tests for the public `appendRichTextRun` immutable append helper.

import { describe, expect, it } from 'vitest';
import { appendRichTextRun, richText } from '../../src/cell';

describe('appendRichTextRun', () => {
  it('appends a plain run, growing 1-run RichText to 2 runs', () => {
    const rt = richText('hello ');
    const out = appendRichTextRun(rt, 'world');
    expect(out.length).toBe(2);
    expect(out[0]?.text).toBe('hello ');
    expect(out[1]?.text).toBe('world');
    expect(out[1]?.font).toBeUndefined();
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('appends a run with the supplied font', () => {
    const rt = richText('plain ');
    const out = appendRichTextRun(rt, 'bold', { b: true });
    expect(out.length).toBe(2);
    expect(out[1]?.text).toBe('bold');
    expect(out[1]?.font).toEqual({ b: true });
  });

  it('does not mutate the source RichText (returns a new frozen array)', () => {
    const rt = richText('x');
    const out = appendRichTextRun(rt, 'y');
    expect(rt.length).toBe(1);
    expect(out.length).toBe(2);
    expect(out).not.toBe(rt);
    expect(Object.isFrozen(rt)).toBe(true);
    expect(Object.isFrozen(out)).toBe(true);
  });
});

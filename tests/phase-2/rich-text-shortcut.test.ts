// Tests for the public `richText(text, font?)` 1-run shortcut.

import { describe, expect, it } from 'vitest';
import { richText } from '../../src/cell';

describe('richText', () => {
  it('produces a frozen 1-run RichText without font when called with text only', () => {
    const rt = richText('hello');
    expect(Array.isArray(rt)).toBe(true);
    expect(Object.isFrozen(rt)).toBe(true);
    expect(rt.length).toBe(1);
    expect(rt[0]?.text).toBe('hello');
    expect(rt[0]?.font).toBeUndefined();
  });

  it('produces a 1-run RichText carrying the supplied font', () => {
    const rt = richText('bold', { b: true, sz: 14 });
    expect(rt.length).toBe(1);
    expect(rt[0]?.text).toBe('bold');
    expect(rt[0]?.font).toEqual({ b: true, sz: 14 });
  });

  it('throws TypeError when text is not a string', () => {
    expect(() => richText(123 as unknown as string)).toThrow(TypeError);
    expect(() => richText(undefined as unknown as string)).toThrow(TypeError);
  });
});

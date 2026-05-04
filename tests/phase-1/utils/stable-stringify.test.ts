import { describe, expect, it } from 'vitest';
import { stableStringify } from '../../../src/utils/stable-stringify';

describe('stableStringify', () => {
  it('matches JSON.stringify for primitives', () => {
    expect(stableStringify(0)).toBe('0');
    expect(stableStringify('hello')).toBe('"hello"');
    expect(stableStringify(null)).toBe('null');
    expect(stableStringify(true)).toBe('true');
  });

  it('sorts object keys recursively', () => {
    expect(stableStringify({ b: 1, a: 2 })).toBe('{"a":2,"b":1}');
    expect(stableStringify({ a: { z: 1, y: 2, x: 3 }, b: 4 })).toBe('{"a":{"x":3,"y":2,"z":1},"b":4}');
  });

  it('preserves array element order', () => {
    expect(stableStringify([3, 1, 2])).toBe('[3,1,2]');
  });

  it('produces identical output for objects with different insertion orders', () => {
    expect(stableStringify({ a: 1, b: 2, c: 3 })).toBe(stableStringify({ c: 3, a: 1, b: 2 }));
  });

  it('throws on circular references (JSON.stringify stack overflow)', () => {
    const a: Record<string, unknown> = {};
    a['self'] = a;
    expect(() => stableStringify(a)).toThrow();
  });
});

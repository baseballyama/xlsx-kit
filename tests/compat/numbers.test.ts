import { describe, expect, it } from 'vitest';
import { isFiniteNumber, isInteger, isTypedArray } from '../../../src/compat/numbers';

describe('isFiniteNumber', () => {
  it('returns true for ordinary finite numbers', () => {
    expect(isFiniteNumber(0)).toBe(true);
    expect(isFiniteNumber(1.5)).toBe(true);
    expect(isFiniteNumber(-1e9)).toBe(true);
    expect(isFiniteNumber(Number.MAX_SAFE_INTEGER)).toBe(true);
  });

  it('returns false for NaN and infinities', () => {
    expect(isFiniteNumber(Number.NaN)).toBe(false);
    expect(isFiniteNumber(Number.POSITIVE_INFINITY)).toBe(false);
    expect(isFiniteNumber(Number.NEGATIVE_INFINITY)).toBe(false);
  });

  it('does NOT coerce strings (unlike global isFinite)', () => {
    expect(isFiniteNumber('1')).toBe(false);
    expect(isFiniteNumber('not a number')).toBe(false);
  });

  it('returns false for non-numbers', () => {
    expect(isFiniteNumber(null)).toBe(false);
    expect(isFiniteNumber(undefined)).toBe(false);
    expect(isFiniteNumber({})).toBe(false);
    expect(isFiniteNumber(true)).toBe(false);
    expect(isFiniteNumber(0n as unknown)).toBe(false);
  });
});

describe('isInteger', () => {
  it('matches Number.isInteger semantics', () => {
    expect(isInteger(0)).toBe(true);
    expect(isInteger(-42)).toBe(true);
    expect(isInteger(1.0)).toBe(true);
    expect(isInteger(1.5)).toBe(false);
    expect(isInteger(Number.NaN)).toBe(false);
    expect(isInteger('1')).toBe(false);
  });
});

describe('isTypedArray', () => {
  it('detects typed-array views', () => {
    expect(isTypedArray(new Uint8Array(2))).toBe(true);
    expect(isTypedArray(new Float64Array(1))).toBe(true);
    expect(isTypedArray(new Int32Array(1))).toBe(true);
  });

  it('rejects DataView and plain ArrayBuffer / arrays', () => {
    expect(isTypedArray(new DataView(new ArrayBuffer(8)))).toBe(false);
    expect(isTypedArray(new ArrayBuffer(8))).toBe(false);
    expect(isTypedArray([])).toBe(false);
    expect(isTypedArray(null)).toBe(false);
  });
});

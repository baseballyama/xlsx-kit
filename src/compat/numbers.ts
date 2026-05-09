// Numeric type guards. The compat layer in the TS port is intentionally tiny —
// most of openpyxl's compat surface (Singleton metaclass, NUMERIC_TYPES tuple,
// py2/py3 shims) has no equivalent in TypeScript.

/**
 * `Number.isFinite` typed as a guard. Distinct from the global `isFinite()`
 * which coerces strings. Rejects `NaN` and ±Infinity.
 */
export function isFiniteNumber(x: unknown): x is number {
  return typeof x === 'number' && Number.isFinite(x);
}

/** `Number.isInteger` typed as a guard. */
export function isInteger(x: unknown): x is number {
  return Number.isInteger(x);
}

/**
 * Detect any of the JS typed-array views (Int8Array, Uint8Array, Float32Array,
 * etc.). Used at workbook-build boundaries where integration code may pass
 * numpy-style buffers — we accept them and iterate per-element rather than
 * refusing.
 */
export function isTypedArray(x: unknown): x is ArrayBufferView {
  return ArrayBuffer.isView(x) && !(x instanceof DataView);
}

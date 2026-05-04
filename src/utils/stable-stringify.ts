// Deterministic JSON serialiser. Used by the Stylesheet pool to dedupe
// value-object entries (Font / Fill / Border / CellXf …) by structural
// equality regardless of property insertion order.
//
// Implemented as a JSON.stringify replacer that returns a key-sorted
// shallow copy of every plain object it encounters; arrays keep their
// element order. Circular references propagate as JSON.stringify's
// native RangeError ("Maximum call stack size exceeded").

const sortKeysReplacer = (_key: string, value: unknown): unknown => {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return value;
  const obj = value as Record<string, unknown>;
  const keys = Object.keys(obj).sort();
  const out: Record<string, unknown> = {};
  for (const k of keys) out[k] = obj[k];
  return out;
};

/**
 * Stringify `value` with object keys sorted recursively. Equal logical
 * values produce the same string regardless of insertion order; arrays
 * preserve their element order. Circular references throw (the error
 * comes straight from JSON.stringify's stack-overflow check).
 */
export function stableStringify(value: unknown): string {
  return JSON.stringify(value, sortKeysReplacer);
}

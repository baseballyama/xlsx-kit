// Inline rich-text runs. Mirrors openpyxl/openpyxl/cell/rich_text.py +
// cell/text.py.
//
// A rich-text cell value is `{ kind: 'rich-text', runs }`. Each run is
// a string segment with an optional InlineFont describing the in-line
// formatting (font name, size, bold / italic / underline, colour, …).
//
// Run-level fields use OOXML's short attribute names (`sz`, `b`, `i`,
// `u`) so the writer can splice them into `<rPr>` directly without
// renaming.

import type { Color } from '../styles/colors';

/** Underline styles per openpyxl's cell-level NestedNoneSet. */
export type InlineUnderline = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export type InlineVertAlign = 'baseline' | 'superscript' | 'subscript';

export interface InlineFont {
  readonly name?: string;
  readonly sz?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  readonly u?: InlineUnderline;
  readonly strike?: boolean;
  readonly outline?: boolean;
  readonly shadow?: boolean;
  readonly condense?: boolean;
  readonly extend?: boolean;
  readonly vertAlign?: InlineVertAlign;
  readonly color?: Color;
  readonly family?: number;
  readonly charset?: number;
  readonly scheme?: 'major' | 'minor';
}

export interface TextRun {
  readonly text: string;
  readonly font?: InlineFont;
}

/** A frozen array of TextRuns. The shared cell value shape under `kind: 'rich-text'`. */
export type RichText = ReadonlyArray<TextRun>;

export function makeTextRun(text: string, font?: InlineFont): TextRun {
  if (typeof text !== 'string') {
    throw new TypeError('makeTextRun: text must be a string');
  }
  return Object.freeze(font !== undefined ? { text, font } : { text });
}

export function makeRichText(runs: ReadonlyArray<TextRun | { text: string; font?: InlineFont }>): RichText {
  const out = runs.map((r) => (Object.isFrozen(r) ? (r as TextRun) : makeTextRun(r.text, r.font)));
  return Object.freeze(out);
}

/** 1-run RichText shortcut. `richText(text, font?)` ≡ `makeRichText([{ text, font }])`. */
export function richText(text: string, font?: InlineFont): RichText {
  if (typeof text !== 'string') {
    throw new TypeError('richText: text must be a string');
  }
  return makeRichText([font !== undefined ? { text, font } : { text }]);
}

/** Return a new frozen RichText with `(text, font?)` appended. The input is not mutated. */
export function appendRichTextRun(rt: RichText, text: string, font?: InlineFont): RichText {
  return makeRichText([...rt, makeTextRun(text, font)]);
}

/**
 * Apply `fn` to each run, returning a new frozen RichText. Useful for run-level
 * bulk transforms (e.g. add `b: true` to every run). The input is not mutated.
 */
export function mapRichTextRuns(
  rt: RichText,
  fn: (run: TextRun, index: number) => TextRun | { text: string; font?: InlineFont },
): RichText {
  return makeRichText(Array.from(rt, fn));
}

/**
 * Apply a common `font` to every run, merging per-run font on top so existing
 * per-run formatting takes precedence. The input is not mutated.
 */
export function applyFontToRichText(rt: RichText, font: InlineFont): RichText {
  return mapRichTextRuns(rt, (r) => ({ text: r.text, font: { ...font, ...(r.font ?? {}) } }));
}

/**
 * Split each run into one run per Unicode code point, preserving the parent
 * run's font on every produced run. Empty-text runs are dropped. Useful as a
 * preprocessing step for per-character styling or animation.
 */
export function splitRichTextRuns(rt: RichText): RichText {
  const out: TextRun[] = [];
  for (const r of rt) {
    if (r.text === '') continue;
    for (const ch of Array.from(r.text)) out.push(makeTextRun(ch, r.font));
  }
  return makeRichText(out);
}

/**
 * Slice the concatenated text of all runs as if it were one string, returning a
 * new RichText covering `[start, end)`. Each output run keeps its parent run's
 * font. Negative indices are interpreted relative to the total length, mirroring
 * `String.prototype.slice`. An out-of-range or empty range yields `[]`.
 */
export function sliceRichText(rt: RichText, start: number, end?: number): RichText {
  const total = richTextLength(rt);
  let s = Math.trunc(start);
  let e = end === undefined ? total : Math.trunc(end);
  if (s < 0) s = Math.max(0, total + s);
  if (e < 0) e = Math.max(0, total + e);
  s = Math.min(s, total);
  e = Math.min(e, total);
  if (s >= e) return makeRichText([]);
  const out: TextRun[] = [];
  let cursor = 0;
  for (const r of rt) {
    const runLen = r.text.length;
    if (runLen === 0) continue;
    const runStart = cursor;
    const runEnd = cursor + runLen;
    cursor = runEnd;
    if (runEnd <= s) continue;
    if (runStart >= e) break;
    const localStart = Math.max(0, s - runStart);
    const localEnd = Math.min(runLen, e - runStart);
    out.push(makeTextRun(r.text.slice(localStart, localEnd), r.font));
  }
  return makeRichText(out);
}

/**
 * Replace `[start, end)` of `rt` with `replacement` (a `RichText` or font-less
 * `string`), returning a new RichText. Negative indices follow
 * `String.prototype.slice`. The input is not mutated.
 */
export function replaceRichText(
  rt: RichText,
  start: number,
  end: number,
  replacement: RichText | string,
): RichText {
  const total = richTextLength(rt);
  let s = Math.trunc(start);
  let e = Math.trunc(end);
  if (s < 0) s = Math.max(0, total + s);
  if (e < 0) e = Math.max(0, total + e);
  s = Math.min(s, total);
  e = Math.min(e, total);
  if (s > e) e = s;
  const before = sliceRichText(rt, 0, s);
  const after = sliceRichText(rt, e, total);
  if (typeof replacement === 'string') {
    return replacement === ''
      ? concatRichText(before, after)
      : concatRichText(before, richText(replacement), after);
  }
  return concatRichText(before, replacement, after);
}

/**
 * Insert `insertion` (a `RichText` or font-less `string`) at `index`.
 * Equivalent to `replaceRichText(rt, index, index, insertion)`. Negative
 * indices follow `String.prototype.slice`.
 */
export function insertRichText(rt: RichText, index: number, insertion: RichText | string): RichText {
  return replaceRichText(rt, index, index, insertion);
}

/**
 * Find the first occurrence of `search` in the concatenated text of `rt`.
 * Mirrors `String.prototype.indexOf` semantics: returns -1 when not found,
 * `fromIndex` defaults to 0, and an empty `search` always matches at
 * `fromIndex` (clamped to `[0, length]`).
 */
export function findRichTextIndex(rt: RichText, search: string, fromIndex?: number): number {
  return richTextToString(rt).indexOf(search, fromIndex);
}

/**
 * Returns true iff the concatenated text of `rt` contains `search`. Mirrors
 * `String.prototype.includes` semantics, including treating an empty
 * `search` as `true`.
 */
export function richTextIncludes(rt: RichText, search: string, fromIndex?: number): boolean {
  return findRichTextIndex(rt, search, fromIndex) >= 0;
}

/**
 * Returns true iff the concatenated text of `rt` starts with `search` at
 * `fromIndex` (default 0). Mirrors `String.prototype.startsWith` semantics,
 * including treating an empty `search` as `true`.
 */
export function richTextStartsWith(rt: RichText, search: string, fromIndex?: number): boolean {
  return richTextToString(rt).startsWith(search, fromIndex);
}

/**
 * Reverse the rich-text by reversing each run's text (code-point-safe) and
 * also reversing the run order. The total concatenated text equals the
 * reverse of `richTextToString(rt)`; per-character font assignments are
 * preserved (each font travels with its character).
 */
export function reverseRichText(rt: RichText): RichText {
  const reversedRuns: TextRun[] = [];
  for (const r of rt) {
    const reversedText = Array.from(r.text).reverse().join('');
    reversedRuns.push(makeTextRun(reversedText, r.font));
  }
  reversedRuns.reverse();
  return makeRichText(reversedRuns);
}

/**
 * Replace every non-overlapping occurrence of `search` with `replacement`,
 * returning a new RichText. An empty `search` is a no-op (returns `rt`
 * unchanged).
 */
export function replaceAllRichText(
  rt: RichText,
  search: string,
  replacement: RichText | string,
): RichText {
  if (search === '') return rt;
  const replacementLen =
    typeof replacement === 'string' ? replacement.length : richTextLength(replacement);
  let cur = rt;
  let from = 0;
  for (;;) {
    const idx = findRichTextIndex(cur, search, from);
    if (idx < 0) break;
    cur = replaceRichText(cur, idx, idx + search.length, replacement);
    from = idx + replacementLen;
  }
  return cur;
}

/**
 * Merge adjacent runs whose `font` is structurally equal, concatenating their
 * `text`. Useful as a cleanup pass after `splitRichTextRuns`, per-char
 * styling, or `concatRichText` chains. The input is not mutated.
 */
export function mergeAdjacentRichTextRuns(rt: RichText): RichText {
  if (rt.length === 0) return makeRichText([]);
  const out: { text: string; font?: InlineFont }[] = [];
  let prevKey: string | undefined;
  for (const r of rt) {
    const key = JSON.stringify(r.font ?? null);
    const last = out[out.length - 1];
    if (last && key === prevKey) {
      last.text += r.text;
    } else {
      out.push(r.font !== undefined ? { text: r.text, font: r.font } : { text: r.text });
      prevKey = key;
    }
  }
  return makeRichText(out);
}

/**
 * Flatten any number of `RichText | string | TextRun` parts into a single
 * frozen RichText. `string` becomes a font-less 1-run; `TextRun` becomes a
 * single run; `RichText` (array) is spread in.
 */
export function concatRichText(...parts: ReadonlyArray<RichText | string | TextRun>): RichText {
  const collected: TextRun[] = [];
  for (const p of parts) {
    if (typeof p === 'string') {
      collected.push(makeTextRun(p));
    } else if (Array.isArray(p)) {
      for (const r of p as ReadonlyArray<TextRun>) collected.push(r);
    } else {
      collected.push(p as TextRun);
    }
  }
  return makeRichText(collected);
}

/**
 * Repeat `rt` `count` times, returning a new RichText. Each run's font is
 * preserved across repetitions. Mirrors `String.prototype.repeat` semantics:
 * `count = 0` (or empty `rt`) yields an empty RichText, `count = 1` returns
 * `rt` unchanged, and a negative or non-finite `count` throws `RangeError`.
 * Fractional `count` is truncated toward zero.
 */
export function repeatRichText(rt: RichText, count: number): RichText {
  if (!Number.isFinite(count) || count < 0) {
    throw new RangeError('repeatRichText: count must be a non-negative finite number');
  }
  const n = Math.floor(count);
  if (n === 0 || rt.length === 0) return makeRichText([]);
  if (n === 1) return rt;
  const parts: RichText[] = new Array(n).fill(rt);
  return concatRichText(...parts);
}

/**
 * Pad the start of `rt` with copies of `padString` (default `' '`) until the
 * concatenated text length reaches `targetLength`, mirroring
 * `String.prototype.padStart`. The pad characters form a single font-less
 * leading run; existing runs (and their fonts) are preserved untouched.
 * Returns `rt` unchanged when `targetLength <= richTextLength(rt)` or when
 * `padString` is empty.
 */
export function padStartRichText(rt: RichText, targetLength: number, padString = ' '): RichText {
  if (padString === '') return rt;
  const cur = richTextLength(rt);
  if (targetLength <= cur) return rt;
  const padded = ''.padStart(targetLength - cur, padString);
  return concatRichText(padded, rt);
}

/**
 * Pad the end of `rt` with copies of `padString` (default `' '`) until the
 * concatenated text length reaches `targetLength`, mirroring
 * `String.prototype.padEnd`. The pad characters form a single font-less
 * trailing run; existing runs (and their fonts) are preserved untouched.
 * Returns `rt` unchanged when `targetLength <= richTextLength(rt)` or when
 * `padString` is empty.
 */
export function padEndRichText(rt: RichText, targetLength: number, padString = ' '): RichText {
  if (padString === '') return rt;
  const cur = richTextLength(rt);
  if (targetLength <= cur) return rt;
  const padded = ''.padEnd(targetLength - cur, padString);
  return concatRichText(rt, padded);
}

/**
 * Trim leading and trailing ASCII whitespace (space, tab, CR, LF) from the
 * concatenated text of `rt`, returning a new RichText. Per-run fonts are
 * preserved on the surviving slice. Internal whitespace is left intact.
 * Returns an empty RichText if every character is whitespace.
 */
export function trimRichText(rt: RichText): RichText {
  const s = richTextToString(rt);
  const total = s.length;
  if (total === 0) return makeRichText([]);
  const firstNon = s.search(/[^ \t\r\n]/);
  if (firstNon < 0) return makeRichText([]);
  let lastNon = total - 1;
  while (lastNon > firstNon) {
    const ch = s.charCodeAt(lastNon);
    if (ch !== 0x20 && ch !== 0x09 && ch !== 0x0d && ch !== 0x0a) break;
    lastNon--;
  }
  return sliceRichText(rt, firstNon, lastNon + 1);
}

/**
 * Concatenate the plain-text content of a rich-text value (rich-text
 * read paths often want the raw text without formatting).
 */
export function richTextToString(rt: RichText): string {
  let out = '';
  for (const r of rt) out += r.text;
  return out;
}

/**
 * Total character count (UTF-16 code units) across all runs.
 * Equivalent to `richTextToString(rt).length` but avoids the string copy.
 */
export function richTextLength(rt: RichText): number {
  let n = 0;
  for (const r of rt) n += r.text.length;
  return n;
}

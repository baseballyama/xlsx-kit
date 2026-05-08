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

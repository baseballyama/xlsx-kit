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

// DrawingML text.
//
// Text body (`<a:txBody>`) is the universal "rich text" container chart
// elements use for titles, axis labels, legend entries, etc. The model covers
// ECMA-376 §21.1.2 (Text Body), §21.1.2.2 (Text Body Properties), §21.1.2.3
// (List Style), §21.1.2.4 (Text Paragraph), §21.1.2.5 (Run / Break / Field),
// and the Run / Paragraph property element groups.

import type { DmlColorWithMods } from './colors';
import type { EffectsRef } from './effect';
import type { Fill } from './fill';
import type { LineProperties } from './line';

// ---- Common building blocks ------------------------------------------------

export type FontAlign = 'auto' | 't' | 'ctr' | 'base' | 'b';
export type ParagraphAlign = 'l' | 'ctr' | 'r' | 'just' | 'justLow' | 'dist' | 'thaiDist';
export type TextUnderline =
  | 'none'
  | 'words'
  | 'sng'
  | 'dbl'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDbl';
export type TextStrike = 'noStrike' | 'sngStrike' | 'dblStrike';
export type TextCap = 'none' | 'small' | 'all';

export type TextOverflow = 'overflow' | 'ellipsis' | 'clip';
export type TextHorzOverflow = 'overflow' | 'clip';
export type TextVertical = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';
export type TextWrap = 'none' | 'square';
export type TextAnchor = 't' | 'ctr' | 'b' | 'just' | 'dist';

/** `<a:lnSpc>` / `<a:spcBef>` / `<a:spcAft>` — either percent (×1000) or points (×100). */
export type TextSpacing = { kind: 'pct'; val: number } | { kind: 'pts'; val: number };

/** `<a:tab>`. */
export interface TabStop {
  pos: number;
  algn?: 'l' | 'ctr' | 'r' | 'dec';
}

/** `<a:latin>` / `<a:ea>` / `<a:cs>` / `<a:sym>`. */
export interface TextFont {
  typeface: string;
  panose?: string;
  pitchFamily?: number;
  charset?: number;
}

/** `<a:hlinkClick>` / `<a:hlinkMouseOver>`. */
export interface HyperlinkInfo {
  rId?: string;
  invalidUrl?: string;
  action?: string;
  tgtFrame?: string;
  tooltip?: string;
  history?: boolean;
  highlightClick?: boolean;
  endSnd?: boolean;
}

/** `<a:rPr>` / `<a:endParaRPr>` / `<a:defRPr>`. */
export interface RunProperties {
  kumimoji?: boolean;
  lang?: string;
  altLang?: string;
  /** Font size in 1/100ths of a point. */
  sz?: number;
  b?: boolean;
  i?: boolean;
  u?: TextUnderline;
  strike?: TextStrike;
  kern?: number;
  cap?: TextCap;
  spc?: number;
  normalizeH?: boolean;
  /** Baseline shift × 1000 (super/subscript). */
  baseline?: number;
  noProof?: boolean;
  dirty?: boolean;
  err?: boolean;
  smtClean?: boolean;
  smtId?: number;
  bmk?: string;
  /** Right-to-left flag. */
  rtl?: boolean;
  ln?: LineProperties;
  fill?: Fill;
  effects?: EffectsRef;
  highlight?: DmlColorWithMods;
  /** Underline-line: either `'follow'` (`<a:uLnTx/>`) or an explicit LineProperties (`<a:uLn>`). */
  uLn?: 'follow' | LineProperties;
  /** Underline-fill: either `'follow'` (`<a:uFillTx/>`) or an explicit Fill (`<a:uFill>`). */
  uFill?: 'follow' | Fill;
  latin?: TextFont;
  ea?: TextFont;
  cs?: TextFont;
  sym?: TextFont;
  hlinkClick?: HyperlinkInfo;
  hlinkMouseOver?: HyperlinkInfo;
}

/** `<a:buChar>` / `<a:buAutoNum>` / `<a:buBlip>` / `<a:buNone>`. */
export type BulletProperties =
  | { kind: 'none' }
  | { kind: 'char'; char: string }
  | {
      kind: 'autoNum';
      type: string;
      startAt?: number;
    }
  | { kind: 'blip'; embedRId?: string; linkRId?: string };

/** `<a:pPr>`. */
export interface ParagraphProperties {
  marL?: number;
  marR?: number;
  /** Indent level (0..8). */
  lvl?: number;
  indent?: number;
  algn?: ParagraphAlign;
  defTabSz?: number;
  rtl?: boolean;
  eaLnBrk?: boolean;
  fontAlgn?: FontAlign;
  latinLnBrk?: boolean;
  hangingPunct?: boolean;
  lnSpc?: TextSpacing;
  spcBef?: TextSpacing;
  spcAft?: TextSpacing;
  tabLst?: TabStop[];
  defRPr?: RunProperties;
  bullet?: BulletProperties;
  /** Bullet font (`<a:buFont>` / `<a:buFontTx/>`). */
  buFont?: 'follow' | TextFont;
  /** Bullet color (`<a:buClr>` / `<a:buClrTx/>`). */
  buClr?: 'follow' | DmlColorWithMods;
  /** Bullet size: percent of run size, points, or "follow run". */
  buSz?: 'follow' | { kind: 'pct'; val: number } | { kind: 'pts'; val: number };
}

/** `<a:r>` (regular run), `<a:br>` (line break), `<a:fld>` (field). */
export type TextRun =
  | { kind: 'r'; rPr?: RunProperties; t: string }
  | { kind: 'br'; rPr?: RunProperties }
  | { kind: 'fld'; id: string; type?: string; rPr?: RunProperties; pPr?: ParagraphProperties; t?: string };

/** `<a:p>`. */
export interface TextParagraph {
  pPr?: ParagraphProperties;
  runs: TextRun[];
  endParaRPr?: RunProperties;
}

/** `<a:bodyPr>` autofit choice. */
export type AutoFit =
  | { kind: 'noAutofit' }
  | { kind: 'normAutofit'; fontScale?: number; lnSpcReduction?: number }
  | { kind: 'spAutoFit' };

/** `<a:bodyPr>`. */
export interface TextBodyProperties {
  rot?: number;
  spcFirstLastPara?: boolean;
  vertOverflow?: TextOverflow;
  horzOverflow?: TextHorzOverflow;
  vert?: TextVertical;
  wrap?: TextWrap;
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
  numCol?: number;
  spcCol?: number;
  rtlCol?: boolean;
  fromWordArt?: boolean;
  anchor?: TextAnchor;
  anchorCtr?: boolean;
  forceAA?: boolean;
  upright?: boolean;
  compatLnSpc?: boolean;
  autoFit?: AutoFit;
  flatTxZ?: number;
}

/**
 * `<a:lstStyle>` (list / level styles). ECMA-376 §21.1.2.4.12. One
 * ParagraphProperties per indent level (0..8). `defPPr` is the default applied
 * when no level-specific override exists.
 */
export interface TextListStyle {
  defPPr?: ParagraphProperties;
  lvl1pPr?: ParagraphProperties;
  lvl2pPr?: ParagraphProperties;
  lvl3pPr?: ParagraphProperties;
  lvl4pPr?: ParagraphProperties;
  lvl5pPr?: ParagraphProperties;
  lvl6pPr?: ParagraphProperties;
  lvl7pPr?: ParagraphProperties;
  lvl8pPr?: ParagraphProperties;
  lvl9pPr?: ParagraphProperties;
}

export interface TextBody {
  bodyPr: TextBodyProperties;
  lstStyle?: TextListStyle;
  paragraphs: TextParagraph[];
}

// ---- Factories -------------------------------------------------------------

export const makeRunProperties = (opts: Partial<RunProperties> = {}): RunProperties => ({ ...opts });

export const makeRun = (text: string, rPr?: RunProperties): TextRun =>
  rPr ? { kind: 'r', rPr, t: text } : { kind: 'r', t: text };

export const makeBreak = (rPr?: RunProperties): TextRun => (rPr ? { kind: 'br', rPr } : { kind: 'br' });

export const makeParagraph = (
  runs: TextRun[],
  pPr?: ParagraphProperties,
  endParaRPr?: RunProperties,
): TextParagraph => ({
  runs,
  ...(pPr ? { pPr } : {}),
  ...(endParaRPr ? { endParaRPr } : {}),
});

export const makeTextBody = (
  paragraphs: TextParagraph[],
  bodyPr: TextBodyProperties = {},
  lstStyle?: TextListStyle,
): TextBody => ({
  bodyPr,
  ...(lstStyle ? { lstStyle } : {}),
  paragraphs,
});

/** Convenience: build a single-paragraph body with one run. */
export const makeSimpleTextBody = (text: string, rPr?: RunProperties): TextBody =>
  makeTextBody([makeParagraph([makeRun(text, rPr)])]);

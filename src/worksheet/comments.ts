// Legacy comments. Per docs/plan/07-rich-features.md §1.
//
// Legacy comments live in `xl/commentsN.xml` (typed comment list +
// authors index) plus a VML drawing part for the popup balloon shape.
// Stage-1 covers plain-text comments + author dedup; rich-text + VML
// shape preservation are reserved for later iterations (we emit a
// minimal placeholder VML so Excel can re-render).

export interface LegacyComment {
  /** Cell reference — typically a single cell ("A1") but Excel allows ranges. */
  ref: string;
  /** Display name of the comment author. */
  author: string;
  /** Plain-text body. Stage-1 doesn't preserve rich-text formatting. */
  text: string;
}

export function makeLegacyComment(opts: { ref: string; author: string; text: string }): LegacyComment {
  return { ref: opts.ref, author: opts.author, text: opts.text };
}

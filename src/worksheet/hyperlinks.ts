// Worksheet hyperlinks.
//
// External URLs land in `xl/worksheets/_rels/sheetN.xml.rels` (one rel per
// Hyperlink entry, type "...relationships/hyperlink", TargetMode="External").
// Internal jumps (`#'Sheet 2'!A1`) live entirely in the `<hyperlink
// location="..."/>` attribute and don't need a rel.

export interface Hyperlink {
  /** Cell or range the hyperlink covers — "A1" or "A1:B5". */
  ref: string;
  /** External URL or relative target path. Mutually exclusive with `location`-only links. */
  target?: string;
  /** Anchor inside the workbook (e.g. `'Sheet 2'!A1`). */
  location?: string;
  /** Tooltip shown on hover. */
  tooltip?: string;
  /** Visible link text — typically falls back to the referenced cell value. */
  display?: string;
  /** Worksheet-rels rId. Populated on read; assigned by the writer when missing. */
  rId?: string;
}

export function makeHyperlink(opts: Partial<Hyperlink> & { ref: string }): Hyperlink {
  if (opts.ref === undefined || opts.ref.length === 0) {
    throw new Error('Hyperlink: ref is required');
  }
  return {
    ref: opts.ref,
    ...(opts.target !== undefined ? { target: opts.target } : {}),
    ...(opts.location !== undefined ? { location: opts.location } : {}),
    ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
    ...(opts.display !== undefined ? { display: opts.display } : {}),
    ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
  };
}

// ---- Worksheet ergonomic helpers ----------------------------------------

import type { Worksheet } from './worksheet';

const replaceHyperlink = (ws: Worksheet, hl: Hyperlink): Hyperlink => {
  const idx = ws.hyperlinks.findIndex((h) => h.ref === hl.ref);
  if (idx >= 0) ws.hyperlinks[idx] = hl;
  else ws.hyperlinks.push(hl);
  return hl;
};

/**
 * Add an external URL hyperlink to a cell or range. The URL goes into the
 * worksheet rels as a hyperlink relationship; the writer generates an rId on
 * save.
 */
export const addUrlHyperlink = (
  ws: Worksheet,
  ref: string,
  url: string,
  opts: { tooltip?: string; display?: string } = {},
): Hyperlink => {
  return replaceHyperlink(
    ws,
    makeHyperlink({
      ref,
      target: url,
      ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
      ...(opts.display !== undefined ? { display: opts.display } : {}),
    }),
  );
};

/**
 * Add an in-workbook jump hyperlink (e.g. to `'Sheet2'!A1` or a defined-name).
 * No rels entry is written — the location is inline in the `<hyperlink
 * location="…"/>` attribute.
 */
export const addInternalHyperlink = (
  ws: Worksheet,
  ref: string,
  location: string,
  opts: { tooltip?: string; display?: string } = {},
): Hyperlink => {
  return replaceHyperlink(
    ws,
    makeHyperlink({
      ref,
      location,
      ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
      ...(opts.display !== undefined ? { display: opts.display } : {}),
    }),
  );
};

/** `mailto:` shortcut. */
export const addMailtoHyperlink = (
  ws: Worksheet,
  ref: string,
  email: string,
  opts: { subject?: string; tooltip?: string; display?: string } = {},
): Hyperlink => {
  const url = opts.subject
    ? `mailto:${email}?subject=${encodeURIComponent(opts.subject)}`
    : `mailto:${email}`;
  return addUrlHyperlink(ws, ref, url, {
    ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
    ...(opts.display !== undefined ? { display: opts.display } : {}),
  });
};

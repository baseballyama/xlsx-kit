// xl/commentsN.xml read/write. Per docs/plan/07-rich-features.md §1.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import type { LegacyComment } from './comments';
import { makeLegacyComment } from './comments';

const COMMENTS_TAG = `{${SHEET_MAIN_NS}}comments`;
const AUTHORS_TAG = `{${SHEET_MAIN_NS}}authors`;
const AUTHOR_TAG = `{${SHEET_MAIN_NS}}author`;
const COMMENT_LIST_TAG = `{${SHEET_MAIN_NS}}commentList`;
const COMMENT_TAG = `{${SHEET_MAIN_NS}}comment`;
const TEXT_TAG = `{${SHEET_MAIN_NS}}text`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const R_TAG = `{${SHEET_MAIN_NS}}r`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/** Concatenate every `<t>` text node found inside a `<text>` body. */
const collectText = (textEl: XmlNode): string => {
  const direct = findChild(textEl, T_TAG);
  if (direct) return direct.text ?? '';
  let out = '';
  for (const child of textEl.children) {
    if (child.name === R_TAG) {
      const t = findChild(child, T_TAG);
      if (t?.text) out += t.text;
    } else if (child.name === T_TAG && child.text) {
      out += child.text;
    }
  }
  return out;
};

/** Parse a `xl/commentsN.xml` payload into a flat LegacyComment list. */
export function parseCommentsXml(bytes: Uint8Array | string): LegacyComment[] {
  const root = parseXml(bytes);
  if (root.name !== COMMENTS_TAG) {
    throw new OpenXmlSchemaError(`parseCommentsXml: root is "${root.name}", expected comments`);
  }
  const authors: string[] = [];
  const authorsEl = findChild(root, AUTHORS_TAG);
  if (authorsEl) {
    for (const a of findChildren(authorsEl, AUTHOR_TAG)) authors.push(a.text ?? '');
  }
  const out: LegacyComment[] = [];
  const listEl = findChild(root, COMMENT_LIST_TAG);
  if (!listEl) return out;
  for (const c of findChildren(listEl, COMMENT_TAG)) {
    const ref = c.attrs['ref'];
    const authorIdRaw = c.attrs['authorId'];
    if (!ref) throw new OpenXmlSchemaError('parseCommentsXml: <comment> missing @ref');
    const authorId = authorIdRaw !== undefined ? Number.parseInt(authorIdRaw, 10) : 0;
    const author = authors[authorId] ?? '';
    const textEl = findChild(c, TEXT_TAG);
    const text = textEl ? collectText(textEl) : '';
    out.push(makeLegacyComment({ ref, author, text }));
  }
  return out;
}

/**
 * Serialise a LegacyComment array to a `xl/commentsN.xml` payload. Authors
 * are deduped: each unique `author` becomes one `<author>` entry, and
 * comments reference it by index.
 */
export function commentsToBytes(comments: ReadonlyArray<LegacyComment>): Uint8Array {
  return new TextEncoder().encode(serializeComments(comments));
}

export function serializeComments(comments: ReadonlyArray<LegacyComment>): string {
  const authorIndex = new Map<string, number>();
  const authors: string[] = [];
  for (const c of comments) {
    if (!authorIndex.has(c.author)) {
      authorIndex.set(c.author, authors.length);
      authors.push(c.author);
    }
  }
  const parts: string[] = [XML_HEADER, `<comments xmlns="${SHEET_MAIN_NS}"><authors>`];
  for (const a of authors) parts.push(`<author>${escapeText(a)}</author>`);
  parts.push('</authors><commentList>');
  for (const c of comments) {
    const id = authorIndex.get(c.author) ?? 0;
    parts.push(
      `<comment ref="${escapeAttr(c.ref)}" authorId="${id}"><text><t>${escapeText(c.text)}</t></text></comment>`,
    );
  }
  parts.push('</commentList></comments>');
  return parts.join('');
}

/**
 * Bare-bones VML drawing payload Excel tolerates as a comment-shape
 * placeholder. We don't preserve the original VML shapes (stage-1 trade-
 * off); this stub guarantees the worksheet rels stay consistent.
 *
 * Carries a single `<v:shape>` with `<x:ClientData ObjectType="Note">`
 * so the load-side content sniffer (src/public/load.ts) classifies the
 * file as comment VML on a re-load — without that marker, the second
 * load → save cycle would mis-classify the placeholder as form-control
 * VML and capture it as passthrough, double-emitting the entry.
 */
export function placeholderVmlDrawing(): Uint8Array {
  return new TextEncoder().encode(
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
      '<v:shape style="visibility:hidden"><x:ClientData ObjectType="Note"/></v:shape>' +
      '</xml>',
  );
}

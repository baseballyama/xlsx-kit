// Streaming XML emitter. Used by writers that need to compose larger documents
// incrementally — chiefly the worksheet writer where the `<sheetData>` block is
// emitted row-by-row to keep heap use bounded.
//
// Phase 1 §5: buffered output only (a chunk array materialised on `result()`).
// Streaming via `WritableStream<Uint8Array>` is added in the phase-4 streaming
// worksheet writer; the structural API stays the same so that change is
// mechanical.
//
// The worksheet hot path emits cells through templated strings, NOT through
// this writer's start / end / writeNode methods. Use `writeRaw` to splice those
// in.

import { escapeXmlAttr, escapeXmlText } from '../utils/escape';
import { OpenXmlIoError } from '../utils/exceptions';
import { utf8ByteLength } from '../utils/utf8';
import { DEFAULT_PREFIXES, parseQName, XML_NS } from './namespaces';
import type { XmlNode } from './tree';

export interface XmlStreamWriterOptions {
  /** Map of namespace URI → prefix. Merged on top of DEFAULT_PREFIXES. */
  prefixMap?: Readonly<Record<string, string>>;
  /** Emit `<?xml … ?>` declaration first. Defaults to true. */
  xmlDeclaration?: boolean;
  /** `standalone` attribute on the declaration. Defaults to 'yes'. */
  standalone?: 'yes' | 'no' | 'omit';
  /**
   * Auto-flush threshold in bytes. Once the in-flight string buffer crosses
   * this size it gets encoded into a chunk and parked. Larger values trade
   * memory for fewer TextEncoder calls; smaller values lower peak memory at the
   * cost of CPU.
   */
  flushBytes?: number;
}

export interface XmlStreamWriter {
  /**
   * Open an element. Names are in Clark notation (`{ns}local`); the writer
   * prefixes them via the configured prefix map.
   */
  start(name: string, attrs?: Record<string, string>): void;
  /** Emit a text node inside the currently open element. */
  text(s: string): void;
  /** Emit a complete subtree. Closes any pending start tag first. */
  writeNode(n: XmlNode): void;
  /** Emit pre-rendered XML bytes verbatim — escape-hatch for hot paths. */
  writeRaw(s: string): void;
  /** Close the currently open element (matched against the start stack). */
  end(): void;
  /** Force any buffered bytes into the chunk store immediately. */
  flush(): void;
  /**
   * Materialise everything written so far. Throws if any element is still open.
   * Idempotent.
   */
  result(): Uint8Array;
}

const DEFAULT_FLUSH_BYTES = 64 * 1024;
const encoder = new TextEncoder();

const buildPrefixMap = (user: Readonly<Record<string, string>> | undefined): Map<string, string> => {
  const out = new Map<string, string>();
  for (const [ns, prefix] of Object.entries(DEFAULT_PREFIXES)) out.set(ns, prefix);
  // The xml prefix is reserved by the XMLNS spec for XML_NS; never override.
  out.set(XML_NS, 'xml');
  if (user !== undefined) {
    for (const [ns, prefix] of Object.entries(user)) out.set(ns, prefix);
  }
  return out;
};

export function createXmlStreamWriter(opts: XmlStreamWriterOptions = {}): XmlStreamWriter {
  const { xmlDeclaration = true, standalone = 'yes', flushBytes = DEFAULT_FLUSH_BYTES } = opts;
  const prefixOf = buildPrefixMap(opts.prefixMap);

  const chunks: Uint8Array[] = [];
  let buf = '';
  // Running UTF-8 byte count of `buf`, maintained incrementally via `append()`.
  // Without this the flush threshold compares UTF-16 code units against a
  // byte budget — non-ASCII payloads then balloon past the configured limit.
  let bufBytes = 0;
  let openStartTag = false;
  let finalised = false;
  const stack: string[] = [];

  const append = (s: string): void => {
    buf += s;
    bufBytes += utf8ByteLength(s);
  };

  const elementName = (name: string): string => {
    const { ns, local } = parseQName(name);
    if (ns === '') return local;
    const prefix = prefixOf.get(ns);
    if (prefix === undefined || prefix === '') return local;
    return `${prefix}:${local}`;
  };
  const attributeName = (name: string): string => {
    const { ns, local } = parseQName(name);
    if (ns === '') return local;
    const prefix = prefixOf.get(ns);
    if (prefix === undefined || prefix === '') return local;
    return `${prefix}:${local}`;
  };

  const flushImpl = (): void => {
    if (buf.length === 0) return;
    chunks.push(encoder.encode(buf));
    buf = '';
    bufBytes = 0;
  };

  const maybeFlush = (): void => {
    if (bufBytes >= flushBytes) flushImpl();
  };

  const closeStartTagIfOpen = (): void => {
    if (!openStartTag) return;
    append('>');
    openStartTag = false;
  };

  if (xmlDeclaration) {
    append('<?xml version="1.0" encoding="UTF-8"');
    if (standalone !== 'omit') append(` standalone="${standalone}"`);
    append('?>\n');
  }

  // ---- writeNode internals
  // ---------------------------------------------------

  const emitNodeInline = (n: XmlNode): void => {
    const tag = elementName(n.name);
    append(`<${tag}`);
    for (const [name, value] of Object.entries(n.attrs)) {
      append(` ${attributeName(name)}="${escapeXmlAttr(value)}"`);
    }
    const text = n.text;
    const hasText = text !== undefined && text !== '';
    const hasChildren = n.children.length > 0;
    if (!hasText && !hasChildren) {
      append('/>');
      return;
    }
    append('>');
    if (hasText) append(escapeXmlText(text));
    for (const c of n.children) emitNodeInline(c);
    append(`</${tag}>`);
  };

  // ---- public surface
  // --------------------------------------------------------

  return {
    start(name, attrs) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: start() after result()');
      closeStartTagIfOpen();
      const tag = elementName(name);
      append(`<${tag}`);
      if (attrs !== undefined) {
        for (const [k, v] of Object.entries(attrs)) {
          append(` ${attributeName(k)}="${escapeXmlAttr(v)}"`);
        }
      }
      stack.push(tag);
      openStartTag = true;
      maybeFlush();
    },
    text(s) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: text() after result()');
      closeStartTagIfOpen();
      append(escapeXmlText(s));
      maybeFlush();
    },
    writeNode(n) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: writeNode() after result()');
      closeStartTagIfOpen();
      emitNodeInline(n);
      maybeFlush();
    },
    writeRaw(s) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: writeRaw() after result()');
      closeStartTagIfOpen();
      append(s);
      maybeFlush();
    },
    end() {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: end() after result()');
      const tag = stack.pop();
      if (tag === undefined) throw new OpenXmlIoError('XmlStreamWriter: end() with no open element');
      if (openStartTag) {
        append('/>');
        openStartTag = false;
      } else {
        append(`</${tag}>`);
      }
      maybeFlush();
    },
    flush() {
      flushImpl();
    },
    result(): Uint8Array {
      if (stack.length > 0) {
        throw new OpenXmlIoError(`XmlStreamWriter: ${stack.length} unclosed element(s) at result()`);
      }
      flushImpl();
      finalised = true;
      let total = 0;
      for (const c of chunks) total += c.byteLength;
      const out = new Uint8Array(total);
      let off = 0;
      for (const c of chunks) {
        out.set(c, off);
        off += c.byteLength;
      }
      return out;
    },
  };
}

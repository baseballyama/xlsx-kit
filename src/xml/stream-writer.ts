// Streaming XML emitter. Used by writers that need to compose larger
// documents incrementally — chiefly the worksheet writer where the
// `<sheetData>` block is emitted row-by-row to keep heap use bounded.
//
// Phase 1 §5: buffered output only (a chunk array materialised on
// `result()`). Streaming via `WritableStream<Uint8Array>` is added in
// the phase-4 streaming worksheet writer; the structural API stays the
// same so that change is mechanical.
//
// Per docs/plan/01-architecture.md §7.2 the worksheet hot path emits
// cells through templated strings, NOT through this writer's start /
// end / writeNode methods. Use `writeRaw` to splice those in.

import { OpenXmlIoError } from '../utils/exceptions';
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
   * Auto-flush threshold in bytes. Once the in-flight string buffer
   * crosses this size it gets encoded into a chunk and parked. Larger
   * values trade memory for fewer TextEncoder calls; smaller values
   * lower peak memory at the cost of CPU.
   */
  flushBytes?: number;
}

export interface XmlStreamWriter {
  /**
   * Open an element. Names are in Clark notation (`{ns}local`); the
   * writer prefixes them via the configured prefix map.
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
   * Materialise everything written so far. Throws if any element is
   * still open. Idempotent.
   */
  result(): Uint8Array;
}

const DEFAULT_FLUSH_BYTES = 64 * 1024;
const encoder = new TextEncoder();

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\r/g, '&#13;')
    .replace(/\n/g, '&#10;')
    .replace(/\t/g, '&#9;');

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
  let openStartTag = false;
  let finalised = false;
  const stack: string[] = [];

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
  };

  const maybeFlush = (): void => {
    if (buf.length >= flushBytes) flushImpl();
  };

  const closeStartTagIfOpen = (): void => {
    if (!openStartTag) return;
    buf += '>';
    openStartTag = false;
  };

  if (xmlDeclaration) {
    buf += '<?xml version="1.0" encoding="UTF-8"';
    if (standalone !== 'omit') buf += ` standalone="${standalone}"`;
    buf += '?>\n';
  }

  // ---- writeNode internals ---------------------------------------------------

  const emitNodeInline = (n: XmlNode): void => {
    const tag = elementName(n.name);
    buf += `<${tag}`;
    for (const [name, value] of Object.entries(n.attrs)) {
      buf += ` ${attributeName(name)}="${escapeAttr(value)}"`;
    }
    const text = n.text;
    const hasText = text !== undefined && text !== '';
    const hasChildren = n.children.length > 0;
    if (!hasText && !hasChildren) {
      buf += '/>';
      return;
    }
    buf += '>';
    if (hasText) buf += escapeText(text);
    for (const c of n.children) emitNodeInline(c);
    buf += `</${tag}>`;
  };

  // ---- public surface --------------------------------------------------------

  return {
    start(name, attrs) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: start() after result()');
      closeStartTagIfOpen();
      const tag = elementName(name);
      buf += `<${tag}`;
      if (attrs !== undefined) {
        for (const [k, v] of Object.entries(attrs)) {
          buf += ` ${attributeName(k)}="${escapeAttr(v)}"`;
        }
      }
      stack.push(tag);
      openStartTag = true;
      maybeFlush();
    },
    text(s) {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: text() after result()');
      closeStartTagIfOpen();
      buf += escapeText(s);
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
      buf += s;
      maybeFlush();
    },
    end() {
      if (finalised) throw new OpenXmlIoError('XmlStreamWriter: end() after result()');
      const tag = stack.pop();
      if (tag === undefined) throw new OpenXmlIoError('XmlStreamWriter: end() with no open element');
      if (openStartTag) {
        buf += '/>';
        openStartTag = false;
      } else {
        buf += `</${tag}>`;
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

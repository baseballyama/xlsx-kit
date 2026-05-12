// SAX iterator over OOXML XML payloads. Wraps `saxes` (XMLNS-aware) and yields
// a flat stream of {start | end | text} events with names already converted to
// Clark notation (`{ns}local`) — same shape as the DOM parser produces for
// static XmlNode trees, so consumers can switch between bulk and streaming
// reads without retouching name comparison.
//
// Phase 1 §3 acceptance: 1 k–row sheetData walked end-to-end with cell counts
// matching the source. The phase-4 read-only worksheet drives real-world use;
// this layer just produces the events.
//
// DOCTYPE / external entity declarations are forbidden. saxes does not expand
// external entities, but a prescan also rejects DTDs in non-streaming inputs.
// Streaming inputs are checked on the first chunk before being fed to the
// parser.

import { SaxesParser } from 'saxes';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { qname } from './namespaces';

export type SaxEvent =
  | { kind: 'start'; name: string; attrs: Record<string, string> }
  | { kind: 'end'; name: string }
  | { kind: 'text'; text: string };

/**
 * Streamable input: `Uint8Array`, plain string, or a Web `ReadableStream` of
 * `Uint8Array` chunks (produced by xlsx zip entries via fflate, fetch, file
 * streams, etc.).
 */
export type SaxInput = Uint8Array | string | ReadableStream<Uint8Array>;

const DOCTYPE_RE = /<!DOCTYPE\b/;
const ENTITY_RE = /<!ENTITY\b/;

const checkDoctype = (text: string): void => {
  if (DOCTYPE_RE.test(text)) {
    throw new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  }
  if (ENTITY_RE.test(text)) {
    throw new OpenXmlSchemaError('Entity declarations are not permitted in OOXML payloads');
  }
};

const isReadableStream = (v: unknown): v is ReadableStream<Uint8Array> => {
  return typeof v === 'object' && v !== null && typeof (v as ReadableStream).getReader === 'function';
};

const decoder = (): TextDecoder => new TextDecoder('utf-8', { fatal: false });

interface SaxesOpenTag {
  name: string;
  uri: string;
  local: string;
  prefix: string;
  attributes: Record<string, { value: string; uri: string; local: string; prefix: string }>;
  isSelfClosing?: boolean;
}

interface SaxesCloseTag {
  name: string;
  uri: string;
  local: string;
  prefix: string;
}

const buildAttrsClark = (attrs: SaxesOpenTag['attributes']): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, info] of Object.entries(attrs)) {
    // saxes already resolved the namespace when xmlns: true is set; raw xmlns /
    // xmlns:* declarations have prefix='xmlns' (or local==='xmlns' when
    // default) and we drop those — they're rebuilt by the serializer.
    if (info.prefix === 'xmlns' || (info.prefix === '' && info.local === 'xmlns')) continue;
    const key = qname(info.uri, info.local);
    out[key] = info.value;
  }
  return out;
};

/**
 * Parse the input as a stream of SAX events. Element / attribute names are
 * returned in Clark notation (`{ns}local`).
 */
export async function* iterParse(input: SaxInput): AsyncIterableIterator<SaxEvent> {
  // Set up the parser. xmlns: true gives us resolved {uri, local, prefix} on
  // every open / close tag and on every attribute.
  const parser = new SaxesParser({ xmlns: true, fragment: false });

  // Head-pointer ring instead of Array#shift: each saxes write() can produce
  // hundreds of events in a single synchronous batch (a `<row>` with dozens of
  // cells flushes one opentag + one text + one closetag per cell). `shift()`
  // is O(n) per element in V8, so a single-batch drain of N events would be
  // O(N²) before iteration. The head advances on yield; the queue is reset
  // (head + length) once it drains so memory stays bounded.
  let queue: SaxEvent[] = [];
  let head = 0;
  let pending: Error | undefined;

  parser.on('error', (err: Error) => {
    pending = err;
  });
  parser.on('doctype', () => {
    pending = new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  });
  parser.on('opentag', (node: SaxesOpenTag) => {
    queue.push({ kind: 'start', name: qname(node.uri, node.local), attrs: buildAttrsClark(node.attributes) });
  });
  parser.on('closetag', (node: SaxesCloseTag) => {
    queue.push({ kind: 'end', name: qname(node.uri, node.local) });
  });
  parser.on('text', (text: string) => {
    if (text.length > 0) queue.push({ kind: 'text', text });
  });

  const drain = function* (): IterableIterator<SaxEvent> {
    for (;;) {
      const ev = head < queue.length ? queue[head] : undefined;
      if (ev === undefined) {
        // Reset rather than grow forever; the next batch starts at index 0.
        queue = [];
        head = 0;
        return;
      }
      head++;
      yield ev;
    }
  };

  const feed = (chunk: string): void => {
    parser.write(chunk);
    if (pending !== undefined) throw pending;
  };

  if (typeof input === 'string') {
    checkDoctype(input);
    feed(input);
  } else if (input instanceof Uint8Array) {
    const text = decoder().decode(input);
    checkDoctype(text);
    feed(text);
    yield* drain();
  } else if (isReadableStream(input)) {
    const reader = input.getReader();
    const td = decoder();
    let firstChunkChecked = false;
    let firstChunkBuffer = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = td.decode(value, { stream: true });
      if (!firstChunkChecked) {
        // We need to see enough of the prologue to be sure no DOCTYPE is
        // hiding. Buffer until we have ~256 chars; if the stream ends before
        // we reach the threshold the tail handler below runs `checkDoctype`
        // on the accumulated prologue. (The previous `|| done` here was dead:
        // a true `done` short-circuits at the top of the loop.)
        firstChunkBuffer += chunk;
        if (firstChunkBuffer.length >= 256) {
          checkDoctype(firstChunkBuffer);
          firstChunkChecked = true;
          feed(firstChunkBuffer);
          yield* drain();
        }
      } else {
        feed(chunk);
        yield* drain();
      }
    }
    // Stream ended; flush decoder + any buffered prologue.
    const tail = td.decode();
    if (!firstChunkChecked) {
      const all = firstChunkBuffer + tail;
      checkDoctype(all);
      feed(all);
      yield* drain();
    } else if (tail.length > 0) {
      feed(tail);
      yield* drain();
    }
  } else {
    throw new OpenXmlSchemaError('iterParse: unsupported input type');
  }

  parser.close();
  if (pending !== undefined) throw pending;
  yield* drain();
}

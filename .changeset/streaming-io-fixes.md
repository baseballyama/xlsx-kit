---
'xlsx-kit': patch
---

Tighten the streaming I/O surface so the README's "fixed-memory" claims hold up
in practice.

- `toFile().toBytes().finish()` no longer re-reads the just-written file. The
  previous code called `fs.readFile(path)` from `finish()` and returned the
  full archive bytes, defeating the chunk-streamed write — a 10M-row workbook
  ended its save by reloading the entire output into memory. `finish()` now
  resolves with an empty `Uint8Array` once the underlying write stream has
  flushed; callers that need the bytes should `fs.readFile()` the path
  themselves.
- `toFile` and `toWritable` honour write-stream backpressure: when `write()`
  returns `false`, subsequent chunks chain off a `drain` event before
  proceeding, so peak memory tracks the writable's `highWaterMark` rather
  than the producer's pace.
- `workbookToBytes` no longer depends on `Buffer`. Browser bundles that omit
  the Node `Buffer` polyfill previously broke at `toBuffer().result()`
  because the in-memory sink ended its result with `Buffer.from(...)`. The
  helper now uses a `Uint8Array`-only sink; a regression test
  (`tests/phase-1/io/browser.test.ts`) saves a workbook with `globalThis.Buffer`
  shadowed to `undefined`.
- The streaming read path inflates worksheet entries chunk-by-chunk. A new
  `ZipArchive.readStream(path)` returns a `ReadableStream<Uint8Array>` that
  drives fflate's `Inflate` incrementally, and `loadWorkbookStream`'s
  whole-sheet `iterRows()` feeds the SAX parser directly off that stream so
  the inflated worksheet body is never fully resident. Band queries
  (`minRow > 1`) still materialise the inflated sheet to build the row-offset
  index — that trade-off is unchanged.
- Documentation: the `XlsxSink` / `BufferedSinkWriter` JSDoc no longer
  describes `toBytes()` as the "buffered mode" — that name was historical;
  the underlying object can either accumulate (buffered sinks) or forward
  chunks as they arrive (streaming sinks). The README also clarifies that
  the streaming reader still loads the compressed archive up front (ZIP
  needs random access to the central directory) — the win is that the
  inflated worksheet payload is never fully resident.

// Write-side I/O abstraction. Symmetric to XlsxSource: a sink is created
// up-front, the writer writes chunks (in either buffered or streaming mode),
// and the caller finalises by calling the corresponding `result*` helper
// exposed on the concrete returned object.

export interface BufferedSinkWriter {
  /** Append a chunk. Must not throw under normal use; errors surface in {@link finish}. */
  write(chunk: Uint8Array): void;
  /**
   * Signal that no more chunks are coming. For in-memory sinks (`toBuffer` /
   * `toBlob` / `toArrayBuffer`) the resolved `Uint8Array` is the full payload;
   * for streaming sinks (`toFile` / `toWritable`) the resolved value is an
   * empty `Uint8Array(0)` once the underlying writable has flushed — the bytes
   * have already been forwarded to disk / the wrapped Writable and are
   * deliberately not buffered for re-emission here. Read the destination via
   * the sink's own `result()` instead.
   */
  finish(): Promise<Uint8Array>;
  /**
   * Abandon the writer. Called by the ZIP writer (and `saveWorkbook` /
   * `loadWorkbookStream` writers) when the surrounding pipeline throws before
   * `finish()` runs. Implementations must release any underlying resource:
   * `toFile` destroys the Node `WriteStream` and best-effort removes the
   * half-written file; `toWritable` destroys the wrapped Writable; in-memory
   * sinks simply drop their buffered chunks.
   *
   * Idempotent. Must not throw — the surrounding catch already has the cause
   * the caller cares about; obscuring it with a cleanup error helps nobody.
   * Optional for backwards compatibility, but writers should provide it.
   */
  abort?(cause?: unknown): void | Promise<void>;
}

export interface XlsxSink {
  /**
   * Chunked-write API. Returns a writer with `write(chunk)` + `finish()`.
   *
   * The name is historical: this is the only entry the ZIP writer drives, and
   * the underlying object can either accumulate chunks in memory (buffered
   * sinks) or forward them to disk / a Writable as they arrive (streaming
   * sinks). The ZIP writer never holds the full archive itself, so the choice
   * of sink decides whether peak memory is "compressed archive size" or
   * "single chunk + deflate scratch".
   *
   * Required: `createZipWriter` calls this on every sink. The method stayed
   * `?:` historically while a `toStream()` alternative was being prototyped,
   * but the streaming-WritableStream variant never landed and leaving it
   * optional only masked the runtime "sink does not expose toBytes()" error
   * behind a missing-property exception.
   */
  toBytes(): BufferedSinkWriter;

  /**
   * Streaming mode: returns a WritableStream into which the writer pipes
   * chunks. Backpressure is honoured by the underlying sink. Currently
   * unused by the ZIP writer (see {@link toBytes}); kept optional so that
   * future writer paths can opt into a `WritableStream`-shaped backend
   * without breaking existing sink implementations.
   */
  toStream?(): WritableStream<Uint8Array>;
}

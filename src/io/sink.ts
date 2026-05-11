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
   */
  toBytes?(): BufferedSinkWriter;

  /**
   * Streaming mode: returns a WritableStream into which the writer pipes
   * chunks. Backpressure is honoured by the underlying sink.
   */
  toStream?(): WritableStream<Uint8Array>;
}

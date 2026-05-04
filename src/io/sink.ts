// Write-side I/O abstraction. Symmetric to XlsxSource: a sink is created
// up-front, the writer writes chunks (in either buffered or streaming
// mode), and the caller finalises by calling the corresponding `result*`
// helper exposed on the concrete returned object.
//
// Per docs/plan/03-foundations.md §1.1.

export interface BufferedSinkWriter {
  /** Append a chunk. Must not throw under normal use; errors surface in {@link finish}. */
  write(chunk: Uint8Array): void;
  /** Resolve all written bytes as a single Uint8Array. */
  finish(): Promise<Uint8Array>;
}

export interface XlsxSink {
  /**
   * Buffered mode: returns an object that accumulates chunks in memory
   * and yields the full payload from {@link BufferedSinkWriter.finish}.
   */
  toBytes?(): BufferedSinkWriter;

  /**
   * Streaming mode: returns a WritableStream into which the writer pipes
   * chunks. Backpressure is honoured by the underlying sink.
   */
  toStream?(): WritableStream<Uint8Array>;
}

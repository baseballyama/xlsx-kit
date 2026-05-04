// Read-side I/O abstraction. The xlsx package layer (src/zip) consumes
// XlsxSource without caring whether the underlying bytes came from a Node
// file handle, a fetch Response, an in-memory Uint8Array or a Blob.
//
// Per docs/plan/03-foundations.md §1.1.

export interface XlsxSource {
  /**
   * Resolve the full payload as a single Uint8Array. Memory-bounded;
   * acceptable for all xlsx parts up to ~hundreds of MB. For larger
   * payloads use {@link toStream}.
   */
  toBytes(): Promise<Uint8Array>;

  /**
   * Sequential byte stream. Optional — implementations that have no
   * cheap streaming representation may omit it. Consumers should
   * fall back to {@link toBytes} when undefined.
   */
  toStream?(): ReadableStream<Uint8Array>;
}

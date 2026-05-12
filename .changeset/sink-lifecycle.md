---
'xlsx-kit': minor
---

fix: release streaming sinks when serialization fails, and bound the streaming write-only buffer against CJK payloads.

When `saveWorkbook` or `createWriteOnlyWorkbook().finalize()` threw partway through serialization, the underlying sink stayed open — `toFile` would leave a half-written `.xlsx` on disk that callers could mistake for a successful save. `BufferedSinkWriter` now exposes an optional `abort(cause?)` hook; the ZIP writer + workbook writers call it from a surrounding `catch` so streaming destinations (`toFile` / `toWritable`) are released and the partial file is best-effort removed. The write-only workbook also exposes its own `abort(cause?)` for callers driving it from a custom pipeline.

The streaming write-only worksheet's pending-byte counter and the XML stream writer's flush threshold now use an accurate UTF-8 length (`utf8ByteLength`) instead of `string.length`. The previous accounting undercounted CJK text by ~3×, letting the in-flight buffer grow well past the configured flush threshold; Japanese / Chinese workloads now flush at the documented ~64 KB ceiling.

API:

- `BufferedSinkWriter.abort(cause?)` is optional. Custom sinks don't need to implement it, but doing so lets `saveWorkbook` clean up streaming destinations on failure.
- `WriteOnlyWorkbook.abort(cause?)` is the matching escape hatch on the write-only API.
- `XlsxSink.toBytes` is now required at the type level. It was already required in practice — the writer threw at runtime when it was missing — and built-in sinks (`toBuffer` / `toBlob` / `toArrayBuffer` / `toFile` / `toWritable`) already implement it.

---
'xlsx-kit': minor
---

Harden the reader against decompression-bomb attacks and tighten release
hygiene ahead of `1.0`:

- `loadWorkbook` / `loadWorkbookStream` now apply a `decompressionLimits`
  guard by default (per-entry size cap, total archive cap, compression-ratio
  cap). The new `OpenXmlDecompressionBombError` (a subclass of
  `OpenXmlIoError`) is thrown when an archive trips the limit. Pass
  `decompressionLimits: false` to disable, or supply a partial override to
  tighten or loosen specific bounds.
- `saveWorkbook` / `workbookToBytes` now validate sheet titles against
  Excel's rules (1–31 chars, forbidden `: \ / ? * [ ]`, no leading/trailing
  apostrophe, reserved `History`, case-insensitive uniqueness) at save time,
  catching invalid names that were introduced by direct mutation of
  `ws.title` after `addWorksheet`.
- `size-limit` now tracks the minified parse size (no brotli, no gzip) of
  `xlsx-kit/streaming` and `xlsx-kit/io` in addition to the existing
  min+brotli budgets, so transitive bundle growth (e.g. the stylesheet
  writer chunk) is caught at PR time.
- New `SECURITY.md` documents the supported versions, the private security
  advisory reporting process, and `decompressionLimits` recommendations for
  consumers.
- New `CONTRIBUTING.md`, GitHub Issue / PR templates, a
  `template-compliance` workflow, and a project-specific `CLAUDE.md` for
  contributors and AI agents working in the repository.
- `docs/migrate-from-openpyxl.md` realigned to the 0.6.x API surface
  (`iterRows`, `setCellByCoord`, `addWorksheet` returning an empty workbook,
  ZIP64 entry-count support, the current passthrough part list).

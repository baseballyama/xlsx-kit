---
'xlsx-kit': patch
---

fix: harden the ZIP64 read path against decompression bombs.

Previously, archives whose End-of-Central-Directory carried ZIP64 sentinel values fell back to `fflate.unzipSync`, which inflates every entry up front. A crafted ZIP64-shaped xlsx could exhaust memory before the per-entry / archive-total caps ran. The random-access reader now parses ZIP64 EOCD + the Zip64 Extended Information extra field directly, so the existing decompression-bomb guards apply to ZIP64 archives the same way they apply to ZIP32.

`loadWorkbook(decompressionLimits)` defaults are unchanged; only the underlying enforcement path is stricter.

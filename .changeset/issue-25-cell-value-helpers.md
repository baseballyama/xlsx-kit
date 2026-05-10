---
'xlsx-kit': minor
---

Extend `cellValueAsString` (in `xlsx-kit/cell`) with optional
`dateFormat` / `emptyText` overrides and add a sibling
`cellValueAsPrimitive` that maps a `CellValue` to the most natural JS
primitive (`string | number | boolean | Date | null`) without forcing a
single target type. Closes #25.

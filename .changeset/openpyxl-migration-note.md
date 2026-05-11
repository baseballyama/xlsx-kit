---
'xlsx-kit': patch
---

Document the small set of openpyxl → xlsx-kit defaults that differ, in particular that `createWorkbook()` returns an empty workbook with no sheets (unlike `openpyxl.Workbook()` which creates a default `'Sheet'`). Direct ports of openpyxl code that include a `wb.remove(wb.active)` call after `Workbook()` were translating that into a no-op `removeSheet(wb, 'Sheet')` — the new README "Migrating from openpyxl" subsection calls this out alongside the `setCell` and `makeBorder` / `makeSide` equivalents. Closes #62.

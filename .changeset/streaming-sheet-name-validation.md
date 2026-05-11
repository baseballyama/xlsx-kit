---
'xlsx-kit': patch
---

Tighten sheet-title validation on the streaming write path. The `createWriteOnlyWorkbook` `addWorksheet` call now applies the same rules as the buffered `addWorksheet` (no `: \ / ? * [ ]`, no leading / trailing apostrophe, not the reserved name `History`) and rejects duplicate titles case-insensitively so the streaming path can't produce a workbook Excel refuses to open.

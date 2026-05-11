---
'xlsx-kit': patch
---

Treat sheet names as case-insensitive for uniqueness, matching Excel. Previously `addWorksheet(wb, 'Data')` followed by `addWorksheet(wb, 'data')` succeeded locally but produced a workbook Excel and LibreOffice refuse to open. `addWorksheet`, `addChartsheet`, `duplicateSheet`, `renameSheet`, and `pickUniqueSheetTitle` now compare titles case-insensitively. A case-only rename of the same sheet (`renameSheet(wb, 'Data', 'data')`) is allowed.

---
'xlsx-kit': minor
---

fix: tighten three worksheet / cell mutation paths that previously produced silently-bad workbooks.

- `copyRange` / `moveRange` no longer carry `hyperlinkId` / `commentId` across worksheets. Those fields are indexes into the source sheet's `hyperlinks` / `legacyComments` arrays and would point at unrelated records (or out of bounds) on the destination. Same-sheet copy / move still preserves them.

- `setSheetState` refuses to hide the last visible sheet. Excel rejects workbooks with every sheet hidden ("Excel cannot use the object linking and embedding features…"); raising at mutation time keeps the workbook recoverable instead of producing a save Excel will reject.

- `makeTextRun` throws `OpenXmlSchemaError` instead of `TypeError` so every public error path uses the documented `OpenXmlError` subclass hierarchy.

---
'xlsx-kit': patch
---

Rename the chart-internal `NumberFormat` interface to `ChartNumberFormat` and re-export it from `xlsx-kit/chart`. The interface was already part of the public surface through `AxisShared.numFmt` and `DataLabelList.numFmt`, but the type itself was not exported — callers building axis / data-label options had to write the literal inline. The new name also disambiguates from the cell-stylesheet `NumberFormat` exported from `xlsx-kit/styles`, which is a different shape (`{ numFmtId, formatCode }`). Closes #58.

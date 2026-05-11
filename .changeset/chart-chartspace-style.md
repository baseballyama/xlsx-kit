---
'xlsx-kit': minor
---

Expose `style?: number` on `ChartSpace` (and `makeChartSpace`). The serializer
emits `<c:style val="N"/>` (range 1..48) between `<c:roundedCorners>` and
`<c:chart>`, selecting one of Excel's built-in "Chart Styles" gallery presets
— the same single attribute openpyxl writes via `chart.style = N`.

Closes #48.

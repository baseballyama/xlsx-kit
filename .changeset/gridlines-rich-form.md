---
'xlsx-kit': minor
---

`AxisShared.majorGridlines` and `AxisShared.minorGridlines` now accept `boolean | Gridlines` instead of just `boolean`. The `Gridlines` shape carries a `ShapeProperties`, so `<c:majorGridlines><c:spPr><a:ln>…</a:ln></c:spPr></c:majorGridlines>` can be emitted to colour / dash / weight the gridline (e.g. corporate-style light grey `D9D9D9`). The plain `true` form keeps emitting `<c:majorGridlines/>` so all existing call sites stay unchanged. Round-trip through `parseChartXml` is preserved for both forms. Closes #57.

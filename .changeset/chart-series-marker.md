---
'xlsx-kit': minor
---

Expose `marker?: Marker` on `LineSeries` and `ScatterSeries` (with the new
`Marker` / `MarkerSymbol` types). The serializer emits `<c:marker>` between
the series' `<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence, carrying
`<c:symbol>`, `<c:size>`, and an optional nested `<c:spPr>` for marker
fill / line colour — matching openpyxl's `series.marker = Marker(...)`.

Closes #47.

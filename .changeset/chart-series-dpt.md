---
'xlsx-kit': minor
---

Expose per-point `dPt?: DataPoint[]` on `BarSeries` (used by bar / line /
area / pie / doughnut / radar / stock / surface), `ScatterSeries`, and
`BubbleSeries`, with the new `DataPoint` type carrying `idx`,
`invertIfNegative?`, `marker?`, `bubble3D?`, `explosion?`, and `spPr?`.
The serializer emits `<c:dPt>` children between the series'
`<c:marker>`/`<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence — unblocking
per-slice colours on pie / doughnut charts, per-bar colours on single-series
bar charts, and per-point styling on line / scatter / bubble.

Closes #44.

---
'xlsx-kit': minor
---

Add `invertIfNegative?: boolean` and `explosion?: number` to `BarSeries`
(used by bar / line / area / pie / doughnut / radar / stock / surface) and
`invertIfNegative?: boolean` to `BubbleSeries`. The serializer emits
`<c:invertIfNegative>` and `<c:explosion>` between `<c:spPr>` and `<c:dPt>`
per ECMA-376 sequence — unblocking per-series colour inversion on negative
values and pie/doughnut slice explosion at the series level (in addition to
the per-point `DataPoint.explosion`).

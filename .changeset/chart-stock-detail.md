---
'xlsx-kit': minor
---

Extend `StockChart.hiLowLines` and `StockChart.upDownBars` to accept a
detailed object form in addition to the existing boolean flag. The
detailed form lets callers style the lines (`HiLowLines.spPr`) and the
up/down bars (`UpDownBars.gapWidth` + `upBars.spPr` + `downBars.spPr`)
with per-element shape properties.

The boolean form (`hiLowLines: true`) keeps its existing meaning and
output, so existing callers are unaffected. Parser round-trips both
forms, picking the boolean form when no detail is found.

New type exports: `BarFrame`, `HiLowLines`, `UpDownBars`.

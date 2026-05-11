---
'xlsx-kit': minor
---

Add `DateAxis` and `SeriesAxis` types and `dateAx?` / `serAx?` slots on
`PlotArea`. `DateAxis` carries `auto`, `lblOffset`, `baseTimeUnit`,
`majorUnit`, `majorTimeUnit`, `minorUnit`, `minorTimeUnit` on top of the
shared axis surface — unblocking time-series charts (`<c:dateAx>`).
`SeriesAxis` adds `tickLblSkip` and `tickMarkSkip`, used by surface charts
(`<c:serAx>`). The serializer emits both inside `<c:plotArea>` between the
inferred cat/val axes and `<c:spPr>`; the parser round-trips them.

New type exports: `DateAxis`, `SeriesAxis`, `TimeUnit`.

---
'xlsx-kit': minor
---

Add `Layout` / `ManualLayout` types and expose `layout?: Layout` on
`ChartTitle`, `PlotArea`, and `Legend`. The serializer emits
`<c:layout><c:manualLayout>` with `layoutTarget`, `xMode` / `yMode` /
`wMode` / `hMode`, and `x` / `y` / `w` / `h` when set, falling back to the
existing empty `<c:layout/>` placeholder when unset — so output is unchanged
for charts that don't configure manual layout. Parser round-trips both
forms.

New type exports: `Layout`, `LayoutMode`, `LayoutTarget`, `ManualLayout`.

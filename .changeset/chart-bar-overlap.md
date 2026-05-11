---
'xlsx-kit': minor
---

Expose `overlap?: number` on `BarChart` (and `makeBarChart`). The serializer
now emits `<c:overlap val="N"/>` (range -100..100) inside `<c:barChart>` when
set, unblocking flush stacking (`overlap: 100`) and negative-space clustered
bars. When unset, the serializer continues to emit the prior default of
`<c:overlap val="100"/>` for `stacked` / `percentStacked` grouping so existing
output is unchanged.

Closes #45.

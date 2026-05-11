---
'xlsx-kit': minor
---

Expose the full ECMA-376 axis attribute surface on `CategoryAxis` and
`ValueAxis`. Previously the serializer emitted fixed defaults for several
elements; these are now driven by typed fields, unblocking horizontal-bar
reversal (`scaling.orientation: 'maxMin'`), 100 %-stacked axis caps
(`scaling.max`), value-axis crossing rules, custom tick formatting, axis
titles, and more.

Newly exposed shared fields: `scaling` (`orientation`/`min`/`max`/`logBase`),
`crosses`, `crossesAt`, `numFmt`, `majorTickMark`, `minorTickMark`,
`tickLblPos`, `title`, `minorGridlines`. `ValueAxis` gains `crossBetween`,
`majorUnit`, `minorUnit`. `CategoryAxis` gains `auto`, `lblAlgn`,
`lblOffset`, `noMultiLvlLbl`. All previously-emitted defaults remain the
output when fields are unset, so existing files are unchanged.

New type exports: `AxisCrossBetween`, `AxisCrosses`, `AxisOrientation`,
`AxisScaling`, `CategoryLabelAlignment`, `TickLabelPosition`, `TickMark`.

Closes #46.

---
'xlsx-kit': minor
---

Add `view3D?: View3D` and `floor?` / `sideWall?` / `backWall?` (typed
`SurfaceFrame`) to `ChartSpace` (and `makeChartSpace`). The serializer
emits `<c:view3D>` (with `rotX`, `rotY`, `depthPercent`, `hPercent`,
`rAngAx`, `perspective`) and `<c:floor>` / `<c:sideWall>` / `<c:backWall>`
(with `thickness` and `spPr`) between `<c:autoTitleDeleted>` and
`<c:plotArea>` per ECMA-376 sequence — unblocking real 3-D chart viewpoints
and wall styling for `bar3DChart` / `line3DChart` / `pie3DChart` /
`area3DChart` / `surface3DChart`.

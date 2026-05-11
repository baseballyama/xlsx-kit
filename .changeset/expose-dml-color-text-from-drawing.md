---
'xlsx-kit': minor
---

Re-export DML colour, fill, and text-body primitives from `xlsx-kit/drawing`. Chart styling reaches `<a:srgbClr>` / `<a:solidFill>` / `<a:bodyPr>…<a:p>…<a:r>` through `ShapeProperties.fill`, `Series.spPr`, `Axis.txPr`, etc., but the building blocks (`DmlColor`, `DmlColorWithMods`, `Fill`, `TextBody`, `TextParagraph`, `RunProperties`, …) and their constructors (`makeColor`, `makeSrgbColor`, `makeSchemeColor`, `makeSolidFill`, `makeTextBody`, `makeParagraph`, `makeRun`, `makeRunProperties`, …) previously had no public home. They now ship as part of `xlsx-kit/drawing` alongside `makeShapeProperties`. Closes #55, closes #56.

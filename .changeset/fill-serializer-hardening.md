---
'xlsx-kit': patch
---

Harden DrawingML `Fill` serializer against two natural mis-uses. (1) Passing a colour without `mods` (e.g. `{ base: { kind: 'srgb', value: 'FF0000' } }` instead of `{ base, mods: [] }`) no longer crashes the chart serializer with `Cannot read properties of undefined (reading 'map')`; the missing modifier list is now treated as empty. (2) Passing a `Fill` with an unknown `kind` (e.g. `'solid'` instead of `'solidFill'`) used to silently emit an empty `<c:spPr></c:spPr>` and lose the caller's styling intent; the serializer now throws `OpenXmlSchemaError` so the mistake surfaces immediately.

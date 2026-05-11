---
'xlsx-kit': patch
---

`addAutoFilterColumn`, `makeHyperlink`, and `setPrintTitles` now throw
`OpenXmlSchemaError` instead of the generic `Error` when their preconditions
are violated. Existing `catch (err instanceof OpenXmlError)` paths now match
these errors uniformly with the rest of the library.

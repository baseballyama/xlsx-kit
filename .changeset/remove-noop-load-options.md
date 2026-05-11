---
'xlsx-kit': minor
---

Remove the silently-ignored `readOnly` / `keepLinks` / `keepVba` / `dataOnly` / `richText` placeholders from `LoadOptions`. They were declared on the public surface but the loader (`src/io/load.ts`) accepted them via `_opts` and dropped them on the floor, so production callers expecting `dataOnly: true` to suppress formulas — or `readOnly: true` to enable a special path — got the default behaviour instead. `LoadOptions` is now an empty type until the underlying behaviour ships; future toggles will land here once they actually do something. The `loadWorkbook(source, opts)` signature is unchanged.

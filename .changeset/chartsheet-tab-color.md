---
'xlsx-kit': minor
---

**Breaking**: `Chartsheet.properties.tabColor` replaces `tabColorRgb`.

Worksheets already expose `SheetProperties.tabColor` as a full `Color` (rgb / indexed / theme / auto / tint). Chartsheets carried a stringly-typed `tabColorRgb` instead, which forced callers to special-case the two sheet kinds and silently dropped every non-RGB colour attribute Excel produces.

The new field is a `Color` object so both sheet kinds share one tab-colour model. Migration:

```ts
// Before
cs.properties = { tabColorRgb: 'FF8800' };

// After
cs.properties = { tabColor: { rgb: 'FF8800' } };
```

Reads recover the additional `indexed` / `theme` / `auto` / `tint` attributes that the old shape discarded.

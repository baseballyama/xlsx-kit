---
'xlsx-kit': patch
---

refactor: route every XML writer through one canonical `escapeXmlAttr` / `escapeXmlText` pair in `src/utils/escape.ts`.

Before, twelve files each carried their own near-identical escape regex — `src/io/save.ts` deliberately skipped `>` while the rest escaped it, and three of them also escaped `\r` / `\n` / `\t` via numeric character references. The discrepancies were quiet correctness bugs (attribute values containing `]]>` rendered differently across writers) and a maintenance hazard.

The unified helpers escape `&`, `<`, `>`, and `"`; whitespace bytes stay literal because our parser (`fast-xml-parser`) does not decode numeric character references and would otherwise break the round-trip.

No user-visible behaviour change beyond consistent attribute escaping across every writer.

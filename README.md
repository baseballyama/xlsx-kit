# xlsx-kit

A TypeScript library for reading and writing Excel `.xlsx` workbooks
from Node 22+ and modern browsers, with no runtime dependencies on
Python or Excel. Inspired by [openpyxl](https://openpyxl.readthedocs.io/).

> **Status: pre-1.0 alpha.** The core read / write / streaming pipeline is
> in place and round-trips real-world fixtures (including pivot tables and
> macro-enabled `.xlsm`), but APIs may shift before `1.0`. See
> [`docs/plan/`](docs/plan/) for the design spec.

## Why xlsx-kit?

The JavaScript xlsx ecosystem in 2026 is split between **commercial upsell
tiers** and **stalled open-source projects**. SheetJS Community Edition
deliberately omits styling, charts, images, pivots, conditional formatting,
and data validation on write — those live in [SheetJS Pro][sjs-pro], a paid
tier. ExcelJS is MIT but
[has not had a meaningful release since October 2023][exceljs-discussion]
and its maintainers explicitly call it inactive; the dependency footprint
unpacks to 21.8 MB. excel4node was [archived in 2022][e4n-archive].
xlsx-js-style is frozen at a 2022 SheetJS fork.

xlsx-kit is the third option: an actively-developed, pure-MIT,
TypeScript-first library with no Pro tier and no missing features behind a
paywall.

[sjs-pro]: https://sheetjs.com/pro/
[exceljs-discussion]: https://github.com/exceljs/exceljs/discussions/2987
[e4n-archive]: https://github.com/natergj/excel4node

| Concern                | Other libraries                                                                  | xlsx-kit                                                              |
|------------------------|----------------------------------------------------------------------------------|-----------------------------------------------------------------------|
| TypeScript types       | hand-written `.d.ts` retrofitted (SheetJS) or community typings (xlsx-populate, excel4node) | first-party, written in TS under `exactOptionalPropertyTypes` + `noUncheckedIndexedAccess` |
| Bundle size            | ExcelJS unpacks to 21.8 MB; xlsx ~7.5 MB                                         | full lib ≤120 KB min+brotli (currently ~85 KB); streaming entry ~49 KB |
| Streaming              | SheetJS docs explicitly note the zip central-directory layout prevents true streaming; ExcelJS supports both directions but the lib is heavy | both read iter and write append, with fixed-memory budget for tens of millions of rows |
| Charts (write)         | none in ExcelJS, xlsx-js-style, SheetJS CE; gated behind SheetJS Pro             | 16 legacy `c:` + 8 modern `cx:` chart kinds (Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap) |
| Pivots / VBA / OLE     | ExcelJS drops pivot tables on read ([#261][exceljs-pivot]); others vary           | byte-identical passthrough so Excel 365 still renders parts we don't model |
| Maintenance            | ExcelJS stalled since 2023; excel4node archived 2022; xlsx-js-style frozen 2022; SheetJS npm artifact frozen 2022 (still distributed via private CDN) | active                                                                |
| License                | SheetJS CE strips features for Pro upsell; SheetJS Pro pricing not published      | MIT, single tier, no upsell                                            |
| Conformance            | none of the major libraries validate against ECMA-376 in CI                       | 3-tier validator (OPC structure + ECMA-376 XSD + semantic invariants) gates every CI build, including a fast-check property-based oracle |
| Modules                | monolithic root barrel                                                           | subpath imports — `xlsx-kit/io`, `/streaming`, `/cell`, `/styles`, etc., each independently tree-shakable |

[exceljs-pivot]: https://github.com/exceljs/exceljs/issues/261

### Where each existing library still wins

- **Read simple xlsx in the browser** → [`read-excel-file`][rexf] is excellent.
- **Write simple xlsx with images** → [`write-excel-file`][wexf] is excellent.
- **Template-based fidelity preservation with password protection** → `xlsx-populate`.
- **Non-xlsx formats** (XLS / XLSB / ODS / CSV / HTML) → SheetJS Community.
- **Commercial budget + long shopping list** → SheetJS Pro.

[rexf]: https://www.npmjs.com/package/read-excel-file
[wexf]: https://www.npmjs.com/package/write-excel-file

### xlsx-kit's home turf

- You write modern TypeScript and want types that actually behave under
  strict mode (cell values are a discriminated union, not `any`).
- You produce **large** xlsx files (tens of millions of cells) and care
  about heap budget.
- You need **charts**, conditional formatting, data validation, defined
  names, tables, ZIP64 — and want them in MIT.
- You round-trip xlsx files that contain pivot tables, VBA macros,
  threaded comments, Power Query metadata, or customXml — and need them
  preserved byte-for-byte.
- You want **proof** that the bytes you emit are valid OOXML, not "Excel
  happens to open them today."

### When NOT to use xlsx-kit

Honest list:

- **Pre-1.0**: API may shift before 1.0. Pin the version for long-running
  projects.
- **`.xlsx` only**: no `.xls` (BIFF), `.xlsb`, `.ods`, or `.csv`. Use
  SheetJS for those.
- **Node 22+ required**: relies on built-in `Web Streams`, `Blob`, and
  `fetch`. Node 18 / 20 (EOL) are not supported.
- **Browser stress-test history is shorter** than ExcelJS's. If you ship
  to millions of browser users today, run your own benchmark first.
- **Visual QA in Excel 365** is on the human-verification list; the schema
  gate proves spec compliance, not that every chart renders pixel-perfect.

### Motivation

The reasons xlsx-kit exists, written down so future contributors don't
relitigate them:

1. **The reference implementation is in Python.** [openpyxl][openpyxl] has
   spent 15 years collecting Excel / LibreOffice corner cases. xlsx-kit
   consumes its fixture corpus directly (`reference/openpyxl/` is a git
   submodule), so edge cases the Python world solved years ago don't get
   re-discovered painfully in JS.
2. **The 2010-era JS stack is heavy.** Most existing libraries pull in
   `jszip`, `lodash`, `archiver`, `xmlbuilder`, `sax`. In 2026 we have
   `fflate`, `fast-xml-parser`, and `saxes` — the toolchain is an order
   of magnitude lighter. xlsx-kit ships with three runtime dependencies.
3. **TypeScript-first changes the API surface.** A library authored in TS
   under strict-mode flags from day one exposes different ergonomics than
   `.d.ts` typings retrofitted onto an old JS codebase.
4. **"Schema-valid" should be a CI gate, not a vibe.** ECMA-376 is
   downloadable; xmllint is free; vendoring the schemas costs <1 MB. There
   is no good reason a 2026 library shouldn't validate every byte it
   emits against the spec.
5. **No Pro tier.** Charts, pivots passthrough, conditional formatting,
   ZIP64 write — all MIT. Nothing held back.

[openpyxl]: https://openpyxl.readthedocs.io/

## Install

```sh
pnpm add xlsx-kit   # or npm / yarn / bun
```

Requires Node `>=22` for the built-in `Web Streams`, `Blob`, and `fetch`
globals.

## Subpath entries

The package has no root barrel — every export lives behind a section
subpath, so your editor's autocomplete only surfaces what's relevant to
the area you're working in. Each export has exactly one home (no
convenience re-exports).

| Import                 | Use case                                          |
|------------------------|---------------------------------------------------|
| `xlsx-kit/io`           | `loadWorkbook` / `saveWorkbook` / `workbookToBytes` plus byte-level Source/Sink + browser helpers (Blob/Response/Stream) |
| `xlsx-kit/node`         | Node fs glue (`fromFile` / `toFile` / `fromBuffer` / `toBuffer` / `fromReadable` / `toWritable`) |
| `xlsx-kit/streaming`    | Read-only iter (`loadWorkbookStream`) + write-only append (`createWriteOnlyWorkbook`) |
| `xlsx-kit/workbook`     | `createWorkbook`, `addWorksheet`, defined names   |
| `xlsx-kit/worksheet`    | `setCell`, `getCell`, `mergeCells`, tables, …     |
| `xlsx-kit/cell`         | Cell value-model + inline rich text               |
| `xlsx-kit/styles`       | Fonts, fills, borders, alignment, number formats  |
| `xlsx-kit/chart`        | `c:` and `cx:` chart kinds                        |
| `xlsx-kit/chartsheet`   | Standalone chartsheets                            |
| `xlsx-kit/drawing`      | Anchors, images, chart placement                  |

Other subpaths: `xlsx-kit/packaging`, `xlsx-kit/utils`, `xlsx-kit/xml`,
`xlsx-kit/zip`, `xlsx-kit/schema`. All exports are tree-shakable
(`"sideEffects": false`).

Bundle budgets (min + brotli):

- `xlsx-kit/streaming` ≤ 80 KB    (currently ~49 KB)
- `xlsx-kit/io` ≤ 120 KB           (currently ~85 KB)

## Quick examples

### Read + edit + write

```ts
import { loadWorkbook, workbookToBytes } from 'xlsx-kit/io';
import { setCell } from 'xlsx-kit/worksheet';
import { fromBuffer } from 'xlsx-kit/node';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromBuffer(await readFile('input.xlsx')));
const sheet = wb.sheets[0];
if (sheet?.kind === 'worksheet') {
  setCell(sheet.sheet, /* row */ 1, /* col */ 1, 'Hello from xlsx-kit');
}
await writeFile('output.xlsx', await workbookToBytes(wb));
```

### Read directly from disk (Node)

```ts
import { loadWorkbook, saveWorkbook } from 'xlsx-kit/io';
import { fromFile, toFile } from 'xlsx-kit/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// …mutate wb…
await saveWorkbook(wb, toFile('output.xlsx'));
```

### Read directly from a `fetch` response (browser)

```ts
import { fromResponse, loadWorkbook } from 'xlsx-kit/io';

const response = await fetch('/sheet.xlsx');
const wb = await loadWorkbook(fromResponse(response));
```

### Streaming write — millions of rows in a fixed memory budget

```ts
import { createWriteOnlyWorkbook } from 'xlsx-kit/streaming';
import { toFile } from 'xlsx-kit/node';

const sink = toFile('big.xlsx');
const wb = await createWriteOnlyWorkbook(sink);
const ws = await wb.addWorksheet('Data');
ws.setColumnWidth(1, 24); // must precede the first appendRow
for (let r = 0; r < 10_000_000; r++) {
  await ws.appendRow([r, `row-${r}`, r * Math.PI]);
}
await ws.close();
await wb.finalize();
```

The streaming writer pushes each row through deflate as it arrives — a
10M-cell archive runs in well under 100 MB heap, the deflate output streams
to disk chunk-by-chunk.

### Streaming read — iterate huge sheets without loading them

```ts
import { loadWorkbookStream } from 'xlsx-kit/streaming';
import { fromFile } from 'xlsx-kit/node';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const sheet = wb.openWorksheet(wb.sheetNames[0] ?? '');
for await (const row of sheet.iterRows({ minRow: 1, maxRow: 100 })) {
  console.log(row.map((c) => c.value));
}
await wb.close();
```

## What's supported

- ✅ Cell values: number, string (sharedStrings), boolean, error, formulas
  (normal / array / shared / dataTable), inline rich text
- ✅ Styles: Font, Fill, Border, Alignment, Protection, NumberFormat, full
  Stylesheet pool with dedup, named styles + DXF
- ✅ Worksheet rich features: mergedCells, sheetView/freezePanes, columnDims,
  rowDims, hyperlinks, defined names, data validations, autoFilter, Tables,
  legacy comments, conditional formatting
- ✅ Drawings: anchors, images (PNG/JPEG/GIF/BMP/WebP/TIFF/SVG/EMF/WMF) with
  format + dimension auto-detection, picture frames in worksheets and charts
- ✅ Charts: 16 legacy `c:` chart kinds + 8 `cx:` chartex kinds (Sunburst,
  Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap),
  spPr / txPr / dLbls / trendline / errBars wiring, chartsheets, UserShapes
- ✅ Pivot tables / VBA / OLE / threaded comments / external links / Power
  Query metadata / customXml / customUI: byte-identical passthrough so
  Excel 365 still renders parts we don't model. The `<workbook>` body
  extras and per-sheet rels chain are preserved end-to-end.
- ✅ Encrypted xlsx detection (CFB Compound Document magic): clear error
  pointing at `msoffcrypto-tool` for decryption.
- ✅ ZIP64 write: workbooks with > 65 535 entries get a ZIP64 EOCD record +
  locator spliced into the final chunk. Read works too.

## Development

The test suite reads fixtures from the `reference/openpyxl` git submodule, so
clone with submodules (or run `pnpm install`, which auto-inits via the
`prepare` script):

```sh
git clone --recursive https://github.com/baseballyama/xlsx-kit.git
# or, if you already cloned without --recursive:
git submodule update --init --recursive

pnpm install
pnpm typecheck
pnpm lint
pnpm test          # vitest, ~2100 tests
pnpm test:perf     # write-only throughput + heap-budget bench
pnpm build         # tsdown + tsc → dist/
pnpm size          # size-limit guards on each bundle
```

[Nix flake](flake.nix) included — `nix develop` (or [direnv](https://direnv.net/)
with `use flake`) gives a pinned Node 22 + pnpm 10 + Python 3 environment.

## License

MIT — see [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

# xlsxify

A TypeScript library for reading and writing Excel `.xlsx` workbooks
from Node 18+ and modern browsers, with no runtime dependencies on
Python or Excel. Inspired by [openpyxl](https://openpyxl.readthedocs.io/).

> **Status: pre-1.0 alpha.** The core read / write / streaming pipeline is
> in place and round-trips real-world fixtures (including pivot tables and
> macro-enabled `.xlsm`), but APIs may shift before `1.0`. See
> [`docs/plan/`](docs/plan/) for the design spec and
> [`PROGRESS.md`](PROGRESS.md) for what has landed in each phase.

## Install

```sh
pnpm add xlsxify   # or npm / yarn / bun
```

Requires Node `>=18.18` for the built-in `Web Streams`, `Blob`, and `fetch`
globals.

## Subpath entries

The package has no root barrel — every export lives behind a section
subpath, so your editor's autocomplete only surfaces what's relevant to
the area you're working in. Each export has exactly one home (no
convenience re-exports).

| Import                 | Use case                                          |
|------------------------|---------------------------------------------------|
| `xlsxify/io`           | `loadWorkbook` / `saveWorkbook` / `workbookToBytes` plus byte-level Source/Sink + browser helpers (Blob/Response/Stream) |
| `xlsxify/node`         | Node fs glue (`fromFile` / `toFile` / `fromBuffer` / `toBuffer` / `fromReadable` / `toWritable`) |
| `xlsxify/streaming`    | Read-only iter (`loadWorkbookStream`) + write-only append (`createWriteOnlyWorkbook`) |
| `xlsxify/workbook`     | `createWorkbook`, `addWorksheet`, defined names   |
| `xlsxify/worksheet`    | `setCell`, `getCell`, `mergeCells`, tables, …     |
| `xlsxify/cell`         | Cell value-model + inline rich text               |
| `xlsxify/styles`       | Fonts, fills, borders, alignment, number formats  |
| `xlsxify/chart`        | `c:` and `cx:` chart kinds                        |
| `xlsxify/chartsheet`   | Standalone chartsheets                            |
| `xlsxify/drawing`      | Anchors, images, chart placement                  |

Other subpaths: `xlsxify/packaging`, `xlsxify/utils`, `xlsxify/xml`,
`xlsxify/zip`, `xlsxify/schema`. All exports are tree-shakable
(`"sideEffects": false`).

Bundle budgets (min + brotli):

- `xlsxify/streaming` ≤ 80 KB    (currently ~49 KB)
- `xlsxify/io` ≤ 120 KB           (currently ~85 KB)

## Quick examples

### Read + edit + write

```ts
import { loadWorkbook, workbookToBytes } from 'xlsxify/io';
import { setCell } from 'xlsxify/worksheet';
import { fromBuffer } from 'xlsxify/node';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromBuffer(await readFile('input.xlsx')));
const sheet = wb.sheets[0];
if (sheet?.kind === 'worksheet') {
  setCell(sheet.sheet, /* row */ 1, /* col */ 1, 'Hello from xlsxify');
}
await writeFile('output.xlsx', await workbookToBytes(wb));
```

### Read directly from disk (Node)

```ts
import { loadWorkbook, saveWorkbook } from 'xlsxify/io';
import { fromFile, toFile } from 'xlsxify/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// …mutate wb…
await saveWorkbook(wb, toFile('output.xlsx'));
```

### Read directly from a `fetch` response (browser)

```ts
import { fromResponse, loadWorkbook } from 'xlsxify/io';

const response = await fetch('/sheet.xlsx');
const wb = await loadWorkbook(fromResponse(response));
```

### Streaming write — millions of rows in a fixed memory budget

```ts
import { createWriteOnlyWorkbook } from 'xlsxify/streaming';
import { toFile } from 'xlsxify/node';

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
import { loadWorkbookStream } from 'xlsxify/streaming';
import { fromFile } from 'xlsxify/node';

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
git clone --recursive https://github.com/baseballyama/openxml-js.git
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

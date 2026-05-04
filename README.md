# openxml-js

A TypeScript port of [openpyxl](https://openpyxl.readthedocs.io/) — read and
write Excel `.xlsx` workbooks from Node 18+ and modern browsers, with no
runtime dependencies on Python or Excel.

> **Status: pre-1.0 alpha.** The core read / write / streaming pipeline is
> in place and round-trips real-world fixtures (including pivot tables and
> macro-enabled `.xlsm`), but APIs may shift before `1.0`. See
> [`docs/plan/`](docs/plan/) for the design spec and
> [`PROGRESS.md`](PROGRESS.md) for what has landed in each phase.

## Install

```sh
pnpm add openxml-js   # or npm / yarn / bun
```

Requires Node `>=18.18` for the built-in `Web Streams`, `Blob`, and `fetch`
globals.

## Subpath entries

| Import                | Use case                                          |
|-----------------------|---------------------------------------------------|
| `openxml-js`          | Full library (workbook model, charts, drawings).  |
| `openxml-js/streaming`| Read-only iter + write-only append. Browser-safe. |
| `openxml-js/node`     | Filesystem / Readable / Writable + the full lib.  |

Bundle budgets (min + brotli):

- `openxml-js`           ≤ 120 KB   (currently ~78 KB)
- `openxml-js/streaming` ≤ 80 KB    (currently ~47 KB)

## Quick examples

### Read + edit + write (full library)

```ts
import { fromBuffer, loadWorkbook, workbookToBytes, setCell } from 'openxml-js';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromBuffer(await readFile('input.xlsx')));
const sheet = wb.sheets[0]?.sheet;
if (sheet?.title) {
  setCell(sheet, /* row */ 1, /* col */ 1, 'Hello from openxml-js');
}
await writeFile('output.xlsx', await workbookToBytes(wb));
```

### Read directly from disk (Node)

```ts
import { fromFile, loadWorkbook, saveWorkbook, toFile } from 'openxml-js/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// …mutate wb…
await saveWorkbook(wb, toFile('output.xlsx'));
```

### Read directly from a `fetch` response (browser)

```ts
import { fromResponse } from 'openxml-js/streaming';
import { loadWorkbook } from 'openxml-js';

const response = await fetch('/sheet.xlsx');
const wb = await loadWorkbook(fromResponse(response));
```

### Streaming write — millions of rows in a fixed memory budget

```ts
import { createWriteOnlyWorkbook } from 'openxml-js/streaming';
import { toFile } from 'openxml-js/node';

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
import { fromFile } from 'openxml-js/node';
import { loadWorkbookStream } from 'openxml-js/streaming';

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

## What's not (yet)

- Date / Duration cell write (read works as numeric serial).
- Random-access streaming reader for sub-sheet cell ranges (the SAX iter
  API is in place; per-cell random access is buffered).
- ZIP64 write — fflate's writer doesn't emit the ZIP64 EOCD record, so we
  fail-fast on > 65 535 entries. Read works.
- Excel for Mac / LibreOffice / Google Sheets visual QA at scale.

## Development

```sh
pnpm install
pnpm typecheck
pnpm lint
pnpm test          # vitest, ~1100 tests
pnpm test:perf     # write-only throughput + heap-budget bench
pnpm build         # tsdown + tsc → dist/
pnpm size          # size-limit guards on each bundle
```

[Nix flake](flake.nix) included — `nix develop` (or [direnv](https://direnv.net/)
with `use flake`) gives a pinned Node 22 + pnpm 10 + Python 3 environment.

## License

MIT — see [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

# Migrating from openpyxl to openxml-js

`openxml-js` is a TypeScript port of openpyxl. The data model, naming, and
semantics line up closely; the API shape is different by necessity (TypeScript
prefers free functions over methods, and the project banned classes outside
of `Error` subclasses).

This guide walks through the most common openpyxl idioms and their
openxml-js equivalents.

## Loading and saving

```python
# openpyxl
from openpyxl import load_workbook
wb = load_workbook('input.xlsx')
wb.save('output.xlsx')
```

```ts
// openxml-js
import { fromFile, loadWorkbook, saveWorkbook, toFile } from 'openxml-js/node';
const wb = await loadWorkbook(fromFile('input.xlsx'));
await saveWorkbook(wb, toFile('output.xlsx'));
```

The `XlsxSource` / `XlsxSink` abstractions decouple the I/O from the workbook,
so the same `loadWorkbook` works against `fromBuffer`, `fromFile`,
`fromBlob`, `fromResponse`, `fromStream`, and `fromReadable`.

## Cells

| openpyxl | openxml-js |
| --- | --- |
| `ws['A1'] = 42` | `setCellByCoord(ws, 'A1', 42)` |
| `ws.cell(row=1, column=1, value=42)` | `setCell(ws, 1, 1, 42)` |
| `ws['A1'].value` | `ws.rows.get(1)?.get(1)?.value` |
| `ws.iter_rows()` | `iterWorksheetRows(ws)` |
| `Cell(formula='=A1+B1')` | `setFormula(cell, 'A1+B1')` |

Cell values cover the same shapes openpyxl does:

- numbers (`number`)
- strings (string, automatically deduped via the shared-strings table)
- booleans (`boolean`)
- formulas (`{ kind: 'formula', formula, t, ... }`)
- errors (`{ kind: 'error', code: '#REF!' }` etc., via `makeErrorValue`)
- rich text (`{ kind: 'rich-text', runs }` via `makeRichText` / `makeTextRun`)
- dates (`Date`)
- durations (`{ kind: 'duration', ms }` via `makeDurationValue`)

## Styles

openpyxl exposes `cell.font`, `cell.fill`, etc. as direct mutable properties
that the cell points at via a styleId. openxml-js keeps the same model — the
`Stylesheet` is a pool — but exposes it through the `cell-style` bridge:

```ts
import {
  makeColor,
  makeFont,
  makePatternFill,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
} from 'openxml-js';

setCellFont(wb, cell, makeFont({ name: 'Arial', size: 14, bold: true }));
setCellFill(wb, cell, makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }));
setCellNumberFormat(wb, cell, '#,##0.00');
```

The pool dedups; assigning the same Font twice yields the same fontId.

## Worksheets

| openpyxl | openxml-js |
| --- | --- |
| `wb.create_sheet('Data')` | `addWorksheet(wb, 'Data')` |
| `wb.active` | `getActiveSheet(wb)` |
| `wb.sheetnames` | `sheetNames(wb)` |
| `wb['Data']` | `getSheet(wb, 'Data')` |
| `del wb['Data']` | `removeSheet(wb, 'Data')` |
| `ws.merged_cells.ranges` | `getMergedCells(ws)` |
| `ws.merge_cells('A1:B2')` | `mergeCells(ws, 'A1:B2')` |
| `ws.freeze_panes = 'B2'` | `setFreezePanes(ws, 'B2')` |

## Streaming write (write_only=True)

```python
# openpyxl
wb = Workbook(write_only=True)
ws = wb.create_sheet('Data')
for row in source:
    ws.append(row)
wb.save('big.xlsx')
```

```ts
// openxml-js
import { createWriteOnlyWorkbook } from 'openxml-js/streaming';
import { toFile } from 'openxml-js/node';

const wb = await createWriteOnlyWorkbook(toFile('big.xlsx'));
const ws = await wb.addWorksheet('Data');
ws.setColumnWidth(1, 20); // must precede the first appendRow
for (const row of source) {
  await ws.appendRow(row);
}
await ws.close();
await wb.finalize();
```

Notable difference: `setColumnWidth` must be called before the first
`appendRow` because `<cols>` is part of the worksheet header that flushes on
the first row. openpyxl is more lenient about this.

## Streaming read (read_only=True)

```python
# openpyxl
wb = load_workbook('big.xlsx', read_only=True)
ws = wb['Data']
for row in ws.iter_rows():
    ...
wb.close()
```

```ts
// openxml-js
import { fromFile } from 'openxml-js/node';
import { loadWorkbookStream } from 'openxml-js/streaming';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const ws = wb.openWorksheet('Data');
for await (const row of ws.iterRows()) {
  // row is an array of ReadOnlyCell { row, col, value, styleId }
}
await wb.close();
```

`iterRows` accepts `{ minRow, maxRow, minCol, maxCol }` for sub-sheet
iteration; the SAX path stops walking the bytes once it crosses `maxRow`.

## What's preserved verbatim (no model)

openxml-js doesn't re-implement every OOXML schema; instead, parts that
aren't worth modelling round-trip byte-for-byte through the workbook's
`passthrough` map. The following live there:

- `xl/vbaProject.bin`, `xl/vbaProjectSignature.bin`
- `xl/pivotCache/`, `xl/pivotTables/`
- `xl/activeX/`, `xl/embeddings/`, `xl/ctrlProps/`
- `xl/printerSettings/`, `xl/queryTables/`, `xl/slicers/`
- `xl/externalLinks/`, `xl/richData/`, `xl/threadedComments/`
- `xl/timelineCaches/`, `xl/timelines/`
- `xl/workbookCache/` (Power Query metadata)
- `customUI/`, `customXml/`
- Control VML drawings (`xl/drawings/*.vml` excluding comment VML)

If openpyxl preserves it, `openxml-js` does too — but typically without a
typed editing surface.

## What's not yet supported

- Editing pivot tables (passthrough only)
- ZIP64 write (the underlying deflate library doesn't emit a ZIP64 EOCD;
  archives with > 65 535 entries fail-fast with `OpenXmlNotImplementedError`)
- Encrypted xlsx (decrypt with `msoffcrypto-tool` first)

See [`docs/plan/`](plan/) for the full design and [`PROGRESS.md`](../PROGRESS.md)
for the current state.

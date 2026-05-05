# E2E — visual-verification scenarios

`pnpm test:e2e` rebuilds `tests/e2e/output/*.xlsx`. Open each file in
Excel (or LibreOffice / Google Sheets / WPS / Numbers) and compare
against the checklist below. The tests themselves only assert that the
files were generated — the *correctness* of the rendered output is the
human reviewer's job.

```sh
pnpm test:e2e
open tests/e2e/output      # macOS
xdg-open tests/e2e/output  # Linux
explorer tests/e2e/output  # Windows
```

If a file fails to open or shows a recovery dialog, that's a defect to
report. If it opens but renders something unexpected, capture a
screenshot and file an issue against the matching scenario.

## Files & checklist

| File | Sheet(s) | What to verify |
|------|----------|----------------|
| `01-basic-cells.xlsx` | Basic | Numbers (column A), strings (A4..A6 incl. `<&>` / multi-line), booleans + the four standard error codes #DIV/0! / #N/A / #REF! / #VALUE! (column B). Column D shows huge + tiny floats. Column C is intentionally empty. |
| `02-formulas.xlsx` | Formulas | A1..A5 = 10..50. **B1** = SUM(A1:A3) shows 60. **C1:C3** is an array formula `{= A1:A3 * 2}`. **D1:D5** is a shared formula derived from D1=A1+1 (relative-shift). **E1** = IF formula returning "positive". **F1** returns #N/A via NA(). |
| `03-dates-windows.xlsx` | 1900 epoch | Column B shows formatted dates (yyyy-mm-dd). Column D shows durations as [h]:mm:ss (90 min → 1:30:00, 2h15m → 2:15:00, 36h → 36:00:00). |
| `03-dates-mac.xlsx` | 1904 epoch | Same dates but interpreted under the Mac 1904 epoch. Useful for checking macOS Excel — should look identical when viewed in macOS Excel set to 1904, but Windows Excel reading it will show the date1904 flag is honoured. |
| `04-rich-text.xlsx` | Rich | A1 currently shows the runs concatenated as a plain string (current stage-1 behaviour — `RichText` stored as inline `<is>/<t>` flat). Real per-run formatting is a known residual. A2 has the same plain-string equivalent. |
| `05-styles.xlsx` | AllAxes / Fonts / Fills / Borders / Align / NumFmt | Each tab demonstrates one style axis. **AllAxes B2** combines bold red Calibri 14pt + yellow fill + thick black borders + center align + protection.locked=false + #,##0.00 — a single cell touching every axis. |
| `06-named-styles.xlsx` | Named Styles | 23 rows mapping each `BUILTIN_NAMED_STYLES` entry to a sample cell. Note: bridging the named style's `cellStyleXfs` index into a `cellXf` so the cell inherits the appearance is a future polish — Excel falls back to the parent style which may not paint identically until the bridge lands. |
| `07-merged-freeze.xlsx` | Merged / Freeze | "Merged" tab: A1:C1 horizontal merge, A2:A4 vertical merge, E5:G7 3×3 block merge. "Freeze" tab: row 1 + col A frozen, scroll right + down to verify the freeze sticks. |
| `08-hyperlinks.xlsx` | Main / Target | A1 → opens GitHub. A2 → jumps to Target!A1 (in-workbook). A3 → external URL with hover tooltip "Hover me". |
| `09-data-validation.xlsx` | Tasks | Column A dropdown of "Open / In Progress / Done" (warning style). Column B integer 0..100 (stop style). Header row has AutoFilter dropdowns. |
| `10-tables.xlsx` | Sales | A1:D6 is an Excel Table named `tblSales` with TableStyleMedium2 (alternating-row banding). Type `=SUM(tblSales[Quantity])` in any empty cell to verify structured-reference autocomplete. |
| `11-comments.xlsx` | Notes | Red-triangle comment markers on A1 / B2 / C3 with author + body visible on hover. |
| `12-conditional-format.xlsx` | CF | Column A "Score" 1..20 with cellIs >15 → red, top-3 → green-bold, 3-color scale. Column B has a data bar. Column C has the 5-arrows icon set. |
| `13-chart-bar.xlsx` | Data | Q1/Q2/Q3 with values 10/30/20. Clustered column chart titled "Quarterly Sales" anchored at D2, ~480×320 px. |
| `14-multi-sheet.xlsx` | Visible1 / Visible2 / Hidden / VeryHidden / Chart Tab | "Hidden" tab requires right-click → Unhide. "VeryHidden" only via API/VBA. "Chart Tab" is a chartsheet showing a pie chart of {A:30, B:50, C:20}. |
| `15-defined-names.xlsx` | Sheet1 | Name Manager (Formulas → Name Manager) shows `total`, `tax`, sheet-scoped `region`, `_xlnm.Print_Area`, `_xlnm.Print_Titles`. File → Print preview should clip to A1:C5 with row 1 repeated on every page. |
| `16-streaming-large.xlsx` | Generated | ~50,000 rows × 6 cols. Should open in <1 s and scroll smoothly. Bottom row label = `even-50000` or `odd-50000`. File size in OS file manager should be a few MB at most. |
| `17-utf8-edge.xlsx` | 売上 / مبيعات / Resumé | Tab labels render correctly across Japanese / Arabic (RTL) / accented Latin. `売上` A1 has emoji 😀 inside a multi-script string. `Resumé` has a cell at the *maximum* coord XFD1048576 ("corner") — press Ctrl+End to jump there. |
| `18-images.xlsx` | Image | A tiny 4×4 PNG block scaled to 96×96 px anchored at C3. |

## Adding a new scenario

1. Create `tests/e2e/scenarios/NN-name.test.ts`. Use `_helpers.ts`'s
   `writeWorkbook(name, wb)` to emit. For streaming scenarios use
   `toFile(...)` directly.
2. The test body must call `expect(result.bytes).toBeGreaterThan(0)`
   (or equivalent for streaming) so a generation crash fails the
   suite.
3. Add the file + checklist row to this README.
4. Run `pnpm test:e2e` and open the file in Excel for sanity.

The output directory is gitignored (`tests/e2e/output/*.xlsx`) so
re-running the suite never produces a noisy diff.

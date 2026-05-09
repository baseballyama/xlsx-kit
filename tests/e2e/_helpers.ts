// Shared helpers for the E2E scenario tests. Each scenario uses
// `outFile(name)` to resolve a path under `tests/e2e/output/` and
// `writeWorkbook(name, wb)` to dump the workbook there. The output
// directory is gitignored — `pnpm test:e2e` rebuilds it.
//
// `writeWorkbook` also performs a load-back smoke check: after writing,
// it feeds the bytes through `loadWorkbook` and asserts that we can at
// least reach the sheet list. This catches ill-formed XML / ZIP that
// visual Excel inspection might gloss over (Excel often shows a
// recovery dialog rather than refusing outright).

import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Workbook } from '../../src/workbook/index';
import { loadWorkbook, workbookToBytes } from '../../src/io/index';
import { fromBuffer } from '../../src/io/node';

export const OUT_DIR = resolve(__dirname, 'output');

mkdirSync(OUT_DIR, { recursive: true });

const outFile = (name: string): string => resolve(OUT_DIR, name);

export const writeWorkbook = async (name: string, wb: Workbook): Promise<{ path: string; bytes: number }> => {
  const bytes = await workbookToBytes(wb);
  const path = outFile(name);
  writeFileSync(path, bytes);

  let loadError: unknown;
  let loadedSheetCount = -1;
  try {
    const reload = await loadWorkbook(fromBuffer(bytes));
    loadedSheetCount = reload.sheets.length;
  } catch (e) {
    loadError = e;
  }

  if (loadError !== undefined) {
    const msg = loadError instanceof Error ? loadError.message : String(loadError);
    throw new Error(`[e2e] ${name} wrote ${bytes.byteLength} bytes but loadWorkbook re-read failed: ${msg}`);
  }

  process.stderr.write(
    `[e2e] wrote ${name}: ${bytes.byteLength.toLocaleString()} bytes (reload: ${loadedSheetCount} sheets) → ${path}\n`,
  );
  return { path, bytes: bytes.byteLength };
};

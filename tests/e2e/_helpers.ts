// Shared helpers for the E2E scenario tests. Each scenario uses
// `outFile(name)` to resolve a path under `tests/e2e/output/` and
// `writeWorkbook(name, wb)` to dump the workbook there. The output
// directory is gitignored — `pnpm test:e2e` rebuilds it.

import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Workbook } from '../../src/index';
import { workbookToBytes } from '../../src/index';

export const OUT_DIR = resolve(__dirname, 'output');

mkdirSync(OUT_DIR, { recursive: true });

export const outFile = (name: string): string => resolve(OUT_DIR, name);

export const writeWorkbook = async (name: string, wb: Workbook): Promise<{ path: string; bytes: number }> => {
  const bytes = await workbookToBytes(wb);
  const path = outFile(name);
  writeFileSync(path, bytes);
  process.stderr.write(`[e2e] wrote ${name}: ${bytes.byteLength.toLocaleString()} bytes → ${path}\n`);
  return { path, bytes: bytes.byteLength };
};

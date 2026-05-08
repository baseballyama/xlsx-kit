// Browser: parse the xlsx the user just selected via <input type="file">.
// fromBlob is streaming, so the workbook starts parsing while the file
// is still being read.

import { loadWorkbook } from 'ooxml-js/xlsx/io';
import { fromBlob } from 'ooxml-js/io';

export async function loadFromInput(input: HTMLInputElement) {
  const file = input.files?.[0];
  if (!file) return null;
  const wb = await loadWorkbook(fromBlob(file));
  return wb;
}

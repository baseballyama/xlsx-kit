// Cell-value XML escape round-trip. Pins the shared-strings + inline-
// string encode/decode loop against the OOXML escape rules:
//
// - `&`, `<`, `>` must escape in text + attribute positions.
// - `"` must escape in attribute positions.
// - `\t`, `\n`, `\r` are XML 1.0 legal whitespace; they survive
//   without re-encoding in text but should not break the parser.
// - Other C0 control characters round-trip via openpyxl's `_xHHHH_`
//   convention (escapeCellString / unescapeCellString in
//   src/utils/escape.ts).
// - Surrounding whitespace in a cell string forces
//   `xml:space="preserve"` on the `<si>/<t>` element.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

const writeAndRead = async (
  values: ReadonlyArray<string>,
): Promise<string[]> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Sheet1');
  values.forEach((v, i) => setCell(ws, i + 1, 1, v));

  const bytes = await workbookToBytes(wb);
  const wb2 = await loadWorkbook(fromBuffer(bytes));
  const ref0 = wb2.sheets[0];
  if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
  return values.map((_, i) => {
    const cell = ref0.sheet.rows.get(i + 1)?.get(1);
    return typeof cell?.value === 'string' ? cell.value : '';
  });
};

describe('phase-3 — XML escape round-trip for cell strings', () => {
  it('preserves the OOXML predefined entities (& < > " \')', async () => {
    const inputs = [
      'A & B',
      'A < B',
      'A > B',
      'A "B" C',
      "A 'B' C",
      'all together: <a href="x">"&"</a>',
    ];
    expect(await writeAndRead(inputs)).toEqual(inputs);
  });

  it('preserves XML-legal whitespace (tab / lf / cr)', async () => {
    const inputs = ['col1\tcol2', 'line1\nline2', 'line1\r\nline2'];
    const out = await writeAndRead(inputs);
    expect(out).toEqual(inputs);
  });

  it('preserves leading + trailing whitespace via xml:space="preserve"', async () => {
    const inputs = ['  leading', 'trailing  ', '  both  ', '\tindented'];
    expect(await writeAndRead(inputs)).toEqual(inputs);
  });

  it('preserves Excel _xHHHH_-encoded control characters', async () => {
    // BEL (), VT (), FS () — all C0 control chars
    // illegal in XML 1.0 text, openpyxl encodes them as _xHHHH_.
    const inputs = ['cellring', 'celltab', 'cellchunk'];
    expect(await writeAndRead(inputs)).toEqual(inputs);
  });

  it('preserves a literal "_x0007_" string (not double-decoded)', async () => {
    // The `_xHHHH_` form is escaping; a user string that *literally*
    // contains "_x0007_" must NOT decode on read. openpyxl handles this
    // by escaping the literal `_x...` to `_x005F_x...`.
    expect(await writeAndRead(['_x0007_'])).toEqual(['_x0007_']);
  });

  it('preserves multi-line cells and rich punctuation', async () => {
    const inputs = [
      'first line\nsecond line\nthird line',
      'mixed: <tag>value & other</tag>\nnext\twith tab',
    ];
    expect(await writeAndRead(inputs)).toEqual(inputs);
  });
});

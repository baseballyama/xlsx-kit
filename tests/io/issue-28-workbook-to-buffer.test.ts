import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { workbookToBuffer } from '../../src/io/node-save';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getCell, setCell } from '../../src/worksheet/worksheet';

describe('workbookToBuffer (#28)', () => {
  it('returns a Node Buffer', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    const buf = await workbookToBuffer(wb);
    expect(Buffer.isBuffer(buf)).toBe(true);
  });

  it('produces the same bytes as workbookToBytes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 42);
    setCell(ws, 2, 1, true);

    const bufBytes = await workbookToBuffer(wb);
    const u8Bytes = await workbookToBytes(wb);

    expect(bufBytes.byteLength).toBe(u8Bytes.byteLength);
    expect(Array.from(bufBytes)).toEqual(Array.from(u8Bytes));
  });

  it('round-trips through loadWorkbook(fromBuffer(...))', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Roundtrip');
    setCell(ws, 1, 1, 'hello');
    setCell(ws, 2, 1, 7);

    const buf = await workbookToBuffer(wb);
    const wb2 = await loadWorkbook(fromBuffer(buf));
    const ref = wb2.sheets[0];
    if (ref === undefined || ref.kind !== 'worksheet') throw new Error('expected a worksheet');
    const ws2 = ref.sheet;

    expect(getCell(ws2, 1, 1)?.value).toBe('hello');
    expect(getCell(ws2, 2, 1)?.value).toBe(7);
  });
});

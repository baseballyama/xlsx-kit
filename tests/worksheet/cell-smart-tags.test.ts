// Tests for the typed worksheet-level <smartTags> model (per-cell
// smart-tag annotations).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet smartTags round-trip', () => {
  it('preserves nested cellSmartTags / cellSmartTag / cellSmartTagPr', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'Alice');
    ws.smartTags.push({
      ref: 'A1',
      tags: [
        {
          type: 0,
          deleted: false,
          xmlBased: false,
          properties: [
            { key: 'firstName', val: 'Alice' },
            { key: 'role', val: 'Engineer' },
          ],
        },
      ],
    });
    ws.smartTags.push({
      ref: 'A2',
      tags: [
        { type: 1, properties: [] },
        { type: 0, deleted: true, properties: [{ key: 'firstName', val: 'Bob' }] },
      ],
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.smartTags.length).toBe(2);
    expect(ws2.smartTags[0]?.ref).toBe('A1');
    expect(ws2.smartTags[0]?.tags[0]?.type).toBe(0);
    expect(ws2.smartTags[0]?.tags[0]?.properties.length).toBe(2);
    expect(ws2.smartTags[0]?.tags[0]?.properties[0]?.key).toBe('firstName');
    expect(ws2.smartTags[0]?.tags[0]?.properties[0]?.val).toBe('Alice');
    expect(ws2.smartTags[1]?.ref).toBe('A2');
    expect(ws2.smartTags[1]?.tags[0]?.type).toBe(1);
    expect(ws2.smartTags[1]?.tags[1]?.deleted).toBe(true);
  });

  it('emits no <smartTags/> when empty', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');
    expect(ws.smartTags.length).toBe(0);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.smartTags.length).toBe(0);
  });
});
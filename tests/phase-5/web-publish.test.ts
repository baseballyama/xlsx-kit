// Tests for the typed worksheet-level <customProperties> + <webPublishItems>
// models.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  makeWebPublishItem,
  makeWorksheetCustomProperty,
} from '../../src/worksheet/web-publish';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet customProperties round-trip', () => {
  it('preserves name (rId omitted when no rels link is set)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setCell(ws, 1, 1, 1);
    ws.customProperties.push(makeWorksheetCustomProperty({ name: 'sharePointLink' }));
    ws.customProperties.push(makeWorksheetCustomProperty({ name: 'topic' }));

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.customProperties.length).toBe(2);
    expect(ws2.customProperties.map((c) => c.name).sort()).toEqual(['sharePointLink', 'topic']);
  });
});

describe('webPublishItems round-trip', () => {
  it('preserves all required + optional attributes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'W');
    setCell(ws, 1, 1, 1);
    ws.webPublishItems.push(
      makeWebPublishItem({
        id: 1,
        divId: 'wp1',
        sourceType: 'range',
        sourceRef: 'A1:B5',
        destinationFile: 'http://example.com/out.html',
        title: 'My published range',
        autoRepublish: true,
      }),
    );
    ws.webPublishItems.push(
      makeWebPublishItem({
        id: 2,
        divId: 'wp2',
        sourceType: 'sheet',
        destinationFile: 'http://example.com/sheet.html',
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.webPublishItems.length).toBe(2);

    const first = ws2.webPublishItems[0];
    expect(first?.id).toBe(1);
    expect(first?.divId).toBe('wp1');
    expect(first?.sourceType).toBe('range');
    expect(first?.sourceRef).toBe('A1:B5');
    expect(first?.destinationFile).toBe('http://example.com/out.html');
    expect(first?.title).toBe('My published range');
    expect(first?.autoRepublish).toBe(true);

    const second = ws2.webPublishItems[1];
    expect(second?.sourceRef).toBeUndefined();
    expect(second?.title).toBeUndefined();
    expect(second?.autoRepublish).toBeUndefined();
  });

  it('emits no <webPublishItems> when none are set', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.webPublishItems.length).toBe(0);
    expect(ws2.customProperties.length).toBe(0);
  });
});
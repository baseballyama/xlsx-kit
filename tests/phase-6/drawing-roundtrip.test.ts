import { describe, expect, it } from 'vitest';
import { makeAbsoluteAnchor, makeOneCellAnchor, makeTwoCellAnchor } from '../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import type { Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet ↔ drawing wiring round-trip', () => {
  it('omits drawing parts when ws.drawing is unset', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoDraw');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('xl/drawings/drawing');
    expect(txt).not.toContain('<drawing r:id');
  });

  it('round-trips a single sheet drawing with one twoCell-anchored chart placeholder', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Charts');
    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'B2', to: 'F12' }), { rId: 'preserved-but-rewritten' }),
    ]);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.drawing).toBeDefined();
    expect(ws2.drawing?.items.length).toBe(1);
    const a = ws2.drawing?.items[0]?.anchor;
    expect(a?.kind).toBe('twoCell');
    if (a?.kind === 'twoCell') {
      expect(a.from.col).toBe(1);
      expect(a.to.col).toBe(5);
      expect(a.to.row).toBe(11);
    }
    // Stage-1 chart placeholder: the writer emits whatever rId was on
    // the chart reference verbatim. Full chart-part allocation lands
    // when the ChartML model arrives.
    const c = ws2.drawing?.items[0]?.content;
    expect(c?.kind).toBe('chart');
    if (c?.kind === 'chart') {
      expect(c.chart.rId).toBe('preserved-but-rewritten');
    }
  });

  it('preserves multiple anchors with mixed kinds', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Mixed');
    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeOneCellAnchor({ from: 'A1', widthPx: 100, heightPx: 50 })),
      makeChartDrawingItem(makeAbsoluteAnchor({ x: 1000, y: 2000, cx: 3000, cy: 4000 })),
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'D1', to: 'G3', editAs: 'oneCell' })),
    ]);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const items = ws2.drawing?.items ?? [];
    expect(items.length).toBe(3);
    expect(items[0]?.anchor.kind).toBe('oneCell');
    expect(items[1]?.anchor.kind).toBe('absolute');
    expect(items[2]?.anchor.kind).toBe('twoCell');
    if (items[2]?.anchor.kind === 'twoCell') expect(items[2].anchor.editAs).toBe('oneCell');
  });

  it('keeps drawings independent across sheets (workbook-global drawingN counter)', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    a.drawing = makeDrawing([makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B2' }))]);
    b.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'C3' })),
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'D1', to: 'E5' })),
    ]);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws0 = expectSheet(wb2.sheets[0]?.sheet);
    const ws1 = expectSheet(wb2.sheets[1]?.sheet);
    expect(ws0.drawing?.items.length).toBe(1);
    expect(ws1.drawing?.items.length).toBe(2);
  });
});

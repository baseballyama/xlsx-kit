import { describe, expect, it } from 'vitest';
import {
  anchorMarkerFromCellRef,
  makeAbsoluteAnchor,
  makeOneCellAnchor,
  makeTwoCellAnchor,
} from '../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing';
import { drawingToBytes, parseDrawingXml } from '../../src/drawing/drawing-xml';
import { EMU_PER_PIXEL } from '../../src/utils/units';

describe('anchor markers + factory functions', () => {
  it('anchorMarkerFromCellRef converts to 0-based col/row with zero offsets', () => {
    expect(anchorMarkerFromCellRef('A1')).toEqual({ col: 0, colOff: 0, row: 0, rowOff: 0 });
    expect(anchorMarkerFromCellRef('B3')).toEqual({ col: 1, colOff: 0, row: 2, rowOff: 0 });
    expect(anchorMarkerFromCellRef('AA10')).toEqual({ col: 26, colOff: 0, row: 9, rowOff: 0 });
  });

  it('rejects malformed coordinates', () => {
    expect(() => anchorMarkerFromCellRef('1A')).toThrowError(/invalid coordinate/);
    expect(() => anchorMarkerFromCellRef('A0')).toThrowError(/invalid coordinate/);
  });

  it('makeAbsoluteAnchor stores pos + ext in EMU', () => {
    const a = makeAbsoluteAnchor({ x: 100, y: 200, cx: 300, cy: 400 });
    expect(a).toEqual({ kind: 'absolute', pos: { x: 100, y: 200 }, ext: { cx: 300, cy: 400 } });
  });

  it('makeOneCellAnchor accepts pixel extent + cell ref', () => {
    const a = makeOneCellAnchor({ from: 'C5', widthPx: 100, heightPx: 50 });
    expect(a.kind).toBe('oneCell');
    if (a.kind !== 'oneCell') return;
    expect(a.from).toEqual({ col: 2, colOff: 0, row: 4, rowOff: 0 });
    expect(a.ext.cx).toBe(100 * EMU_PER_PIXEL);
    expect(a.ext.cy).toBe(50 * EMU_PER_PIXEL);
  });

  it('makeTwoCellAnchor optionally accepts editAs', () => {
    const a = makeTwoCellAnchor({ from: 'A1', to: 'D5', editAs: 'oneCell' });
    expect(a.kind).toBe('twoCell');
    if (a.kind !== 'twoCell') return;
    expect(a.from.col).toBe(0);
    expect(a.to.col).toBe(3);
    expect(a.editAs).toBe('oneCell');
  });
});

describe('parseDrawingXml + serializeDrawing round-trip', () => {
  it('round-trips a twoCell-anchored chart', () => {
    const drawing = makeDrawing([makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'F10' }), { rId: 'rId1' })]);
    const bytes = drawingToBytes(drawing);
    const back = parseDrawingXml(bytes);
    expect(back.items.length).toBe(1);
    const a = back.items[0]?.anchor;
    expect(a?.kind).toBe('twoCell');
    if (a?.kind === 'twoCell') {
      expect(a.from.col).toBe(0);
      expect(a.to.col).toBe(5);
      expect(a.to.row).toBe(9);
    }
    const c = back.items[0]?.content;
    expect(c?.kind).toBe('chart');
    if (c?.kind === 'chart') {
      expect(c.chart.rId).toBe('rId1');
    }
  });

  it('round-trips a oneCell + an absolute anchor in one drawing', () => {
    const drawing = makeDrawing([
      makeChartDrawingItem(makeOneCellAnchor({ from: 'B2', widthPx: 200, heightPx: 100 }), { rId: 'rId7' }),
      makeChartDrawingItem(makeAbsoluteAnchor({ x: 1000, y: 2000, cx: 3000, cy: 4000 }), { rId: 'rId8' }),
    ]);
    const back = parseDrawingXml(drawingToBytes(drawing));
    expect(back.items.length).toBe(2);
    expect(back.items[0]?.anchor.kind).toBe('oneCell');
    expect(back.items[1]?.anchor.kind).toBe('absolute');
    if (back.items[1]?.anchor.kind === 'absolute') {
      expect(back.items[1].anchor.pos).toEqual({ x: 1000, y: 2000 });
      expect(back.items[1].anchor.ext).toEqual({ cx: 3000, cy: 4000 });
    }
  });

  it('preserves twoCell editAs', () => {
    const drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'C3', editAs: 'absolute' }), { rId: 'rId1' }),
    ]);
    const back = parseDrawingXml(drawingToBytes(drawing));
    expect((back.items[0]?.anchor as { editAs?: string }).editAs).toBe('absolute');
  });

  it('rejects a non-wsDr root', () => {
    expect(() => parseDrawingXml('<foo/>')).toThrowError(/expected wsDr/);
  });
});

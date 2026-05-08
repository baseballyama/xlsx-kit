// Tests for the addImageAt + addChartAt drawing helpers.

import { describe, expect, it } from 'vitest';
import { addChartAt, addImageAt } from '../../src/xlsx/drawing/drawing';
import { loadImage } from '../../src/xlsx/drawing/image';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/xlsx/chart/chart';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

const TINY_BLUE_PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAFklEQVR4nGP8z8DAwMDAxMDA8J+BAQAOAQHv6sTncgAAAABJRU5ErkJggg==';

describe('addImageAt', () => {
  it('lazy-allocates ws.drawing and appends a picture item', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(ws.drawing).toBeUndefined();

    const bytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    const item = addImageAt(ws, 'C3', bytes);
    expect(ws.drawing?.items.length).toBe(1);
    expect(item.content.kind).toBe('picture');
  });

  it('accepts an XlsxImage already loaded', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const bytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    const img = loadImage(bytes);
    addImageAt(ws, 'D5', img, { widthPx: 200, heightPx: 100 });
    expect(ws.drawing?.items.length).toBe(1);
  });

  it('appends multiple images on the same drawing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const bytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    addImageAt(ws, 'A1', bytes);
    addImageAt(ws, 'C5', bytes);
    expect(ws.drawing?.items.length).toBe(2);
  });
});

describe('addChartAt', () => {
  it('lazy-allocates ws.drawing and appends a chart item', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'Q1');
    setCell(ws, 2, 1, 100);

    const space = makeChartSpace({
      title: 'Q sales',
      plotArea: {
        chart: makeBarChart({
          barDir: 'col',
          grouping: 'clustered',
          series: [
            makeBarSeries({
              idx: 0,
              val: { ref: 'A!$B$2:$B$2', cache: [100] },
              cat: { ref: 'A!$A$2:$A$2', cacheKind: 'str', cache: ['Q1'] },
            }),
          ],
        }),
      },
      legend: { position: 'r' },
    });

    const item = addChartAt(ws, 'D3', { space });
    expect(ws.drawing?.items.length).toBe(1);
    expect(item.content.kind).toBe('chart');
  });

  it('image and chart can coexist on the same sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const bytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    addImageAt(ws, 'A1', bytes);

    const space = makeChartSpace({
      title: 'X',
      plotArea: { chart: makeBarChart({ series: [] }) },
      legend: { position: 'r' },
    });
    addChartAt(ws, 'D5', { space });
    expect(ws.drawing?.items.length).toBe(2);
    expect(ws.drawing?.items[0]?.content.kind).toBe('picture');
    expect(ws.drawing?.items[1]?.content.kind).toBe('chart');
  });
});
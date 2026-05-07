// Tests for listImagesOnSheet / listChartsOnSheet / removeAllDrawingItems.

import { describe, expect, it } from 'vitest';
import {
  addChartAt,
  addImageAt,
  listChartsOnSheet,
  listImagesOnSheet,
  removeAllDrawingItems,
} from '../../src/drawing/drawing';
import { loadImage } from '../../src/drawing/image';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

// 1×1 transparent PNG as test bytes.
const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('listImagesOnSheet', () => {
  it('returns only picture items', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const img = loadImage(PNG_1X1);
    addImageAt(ws, 'A1', img);
    addChartAt(ws, 'C1', {});
    addImageAt(ws, 'A5', img);
    expect(listImagesOnSheet(ws).length).toBe(2);
  });

  it('empty when no drawing exists', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(listImagesOnSheet(ws)).toEqual([]);
  });
});

describe('listChartsOnSheet', () => {
  it('returns only chart items', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const img = loadImage(PNG_1X1);
    addImageAt(ws, 'A1', img);
    addChartAt(ws, 'C1', {});
    addChartAt(ws, 'C5', {});
    expect(listChartsOnSheet(ws).length).toBe(2);
  });

  it('empty when no drawing exists', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(listChartsOnSheet(ws)).toEqual([]);
  });
});

describe('removeAllDrawingItems', () => {
  it('drops every item and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addImageAt(ws, 'A1', loadImage(PNG_1X1));
    addChartAt(ws, 'C1', {});
    expect(removeAllDrawingItems(ws)).toBe(2);
    expect(ws.drawing?.items).toEqual([]);
  });

  it('returns 0 when no drawing exists', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllDrawingItems(ws)).toBe(0);
  });
});

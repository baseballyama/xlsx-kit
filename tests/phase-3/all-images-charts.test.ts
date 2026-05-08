// Tests for getAllImages / getAllCharts.

import { describe, expect, it } from 'vitest';
import { addChartAt, addImageAt } from '../../src/xlsx/drawing/drawing';
import { loadImage } from '../../src/xlsx/drawing/image';
import {
  addWorksheet,
  createWorkbook,
  getAllCharts,
  getAllImages,
} from '../../src/xlsx/workbook/workbook';

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('getAllImages', () => {
  it('aggregates images across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    const img = loadImage(PNG_1X1);
    addImageAt(a, 'A1', img);
    addImageAt(b, 'A1', img);
    addChartAt(a, 'C1', {});
    const out = getAllImages(wb);
    expect(out.length).toBe(2);
    expect(out[0]?.sheet.title).toBe('A');
    expect(out[1]?.sheet.title).toBe('B');
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(getAllImages(wb)).toEqual([]);
  });
});

describe('getAllCharts', () => {
  it('aggregates charts across every worksheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    addChartAt(a, 'A1', {});
    addChartAt(b, 'A1', {});
    addImageAt(a, 'C1', loadImage(PNG_1X1));
    expect(getAllCharts(wb).length).toBe(2);
  });

  it('skips worksheets without a drawing', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A'); // no drawing
    expect(getAllCharts(wb)).toEqual([]);
  });
});

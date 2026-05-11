// Tests for removeAllImages / removeAllCharts (kind-specific drawing wipes).

import { describe, expect, it } from 'vitest';
import {
  addChartAt,
  addImageAt,
  listChartsOnSheet,
  listImagesOnSheet,
  removeAllCharts,
  removeAllImages,
} from '../../src/drawing/drawing';
import { loadImage } from '../../src/drawing/image';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('removeAllImages', () => {
  it('removes only pictures, leaving charts', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const img = loadImage(PNG_1X1);
    addImageAt(ws, 'A1', img);
    addChartAt(ws, 'C1', {});
    addImageAt(ws, 'A5', img);
    expect(removeAllImages(ws)).toBe(2);
    expect(listImagesOnSheet(ws).length).toBe(0);
    expect(listChartsOnSheet(ws).length).toBe(1);
  });

  it('returns 0 when sheet has no drawing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllImages(ws)).toBe(0);
  });
});

describe('removeAllCharts', () => {
  it('removes only charts, leaving pictures', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const img = loadImage(PNG_1X1);
    addChartAt(ws, 'C1', {});
    addImageAt(ws, 'A1', img);
    addChartAt(ws, 'C5', {});
    expect(removeAllCharts(ws)).toBe(2);
    expect(listChartsOnSheet(ws).length).toBe(0);
    expect(listImagesOnSheet(ws).length).toBe(1);
  });

  it('returns 0 when sheet has no drawing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllCharts(ws)).toBe(0);
  });
});

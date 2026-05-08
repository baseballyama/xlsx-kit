import { describe, expect, it } from 'vitest';
import { makeAbsoluteAnchor, makeOneCellAnchor, makeTwoCellAnchor } from '../../src/xlsx/drawing/anchor';
import { makeDrawing, makePictureDrawingItem } from '../../src/xlsx/drawing/drawing';
import { drawingToBytes, parseDrawingXml } from '../../src/xlsx/drawing/drawing-xml';
import { detectImageDimensions, detectImageFormat, loadImage, type XlsxImage } from '../../src/xlsx/drawing/image';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

// ---- Test fixtures (synthesised in-memory rather than reading binary files) ----

const makePngBytes = (width: number, height: number): Uint8Array => {
  const out = new Uint8Array(33);
  // 8-byte PNG signature.
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  // IHDR chunk: 4 byte length=13, "IHDR", width (BE), height (BE), 1 byte each: depth, color, compression, filter, interlace.
  out.set([0x00, 0x00, 0x00, 0x0d], 8);
  out.set([0x49, 0x48, 0x44, 0x52], 12);
  const w = new DataView(out.buffer, 16, 4);
  w.setUint32(0, width);
  const h = new DataView(out.buffer, 20, 4);
  h.setUint32(0, height);
  // bit depth + color type + compression + filter + interlace + 4-byte CRC (left zeroed).
  out.set([0x08, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00], 24);
  return out;
};

const makeGifBytes = (width: number, height: number): Uint8Array => {
  const out = new Uint8Array(13);
  out.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61], 0); // "GIF89a"
  out[6] = width & 0xff;
  out[7] = (width >> 8) & 0xff;
  out[8] = height & 0xff;
  out[9] = (height >> 8) & 0xff;
  return out;
};

const makeJpegBytes = (width: number, height: number): Uint8Array => {
  const out = new Uint8Array(20);
  out.set([0xff, 0xd8], 0); // SOI
  out.set([0xff, 0xc0, 0x00, 0x11], 2); // SOF0 + length=17
  out[6] = 0x08; // precision
  out[7] = (height >> 8) & 0xff;
  out[8] = height & 0xff;
  out[9] = (width >> 8) & 0xff;
  out[10] = width & 0xff;
  return out;
};

describe('Image format detection', () => {
  it('detects PNG / JPEG / GIF magic bytes', () => {
    expect(detectImageFormat(makePngBytes(10, 10))).toBe('png');
    expect(detectImageFormat(makeJpegBytes(10, 10))).toBe('jpeg');
    expect(detectImageFormat(makeGifBytes(10, 10))).toBe('gif');
  });

  it('detects BMP / WebP / TIFF / SVG', () => {
    expect(detectImageFormat(new Uint8Array([0x42, 0x4d, 0x00, 0x00]))).toBe('bmp');
    const webp = new Uint8Array(16);
    webp.set([0x52, 0x49, 0x46, 0x46], 0);
    webp.set([0x57, 0x45, 0x42, 0x50], 8);
    expect(detectImageFormat(webp)).toBe('webp');
    expect(detectImageFormat(new Uint8Array([0x49, 0x49, 0x2a, 0x00]))).toBe('tiff');
    const svg = new TextEncoder().encode('<?xml version="1.0"?><svg xmlns="..."></svg>');
    expect(detectImageFormat(svg)).toBe('svg');
  });

  it('returns undefined for an unrecognised payload', () => {
    expect(detectImageFormat(new Uint8Array([0x00, 0x01, 0x02, 0x03]))).toBeUndefined();
  });
});

describe('Image dimension extraction', () => {
  it('reads PNG IHDR width/height (big-endian)', () => {
    expect(detectImageDimensions(makePngBytes(640, 480), 'png')).toEqual({ width: 640, height: 480 });
  });

  it('reads GIF logical screen descriptor (little-endian)', () => {
    expect(detectImageDimensions(makeGifBytes(320, 200), 'gif')).toEqual({ width: 320, height: 200 });
  });

  it('reads JPEG SOF0 width/height', () => {
    expect(detectImageDimensions(makeJpegBytes(800, 600), 'jpeg')).toEqual({ width: 800, height: 600 });
  });
});

describe('loadImage', () => {
  it('infers format + dimensions for a PNG', () => {
    const img = loadImage(makePngBytes(120, 80));
    expect(img.format).toBe('png');
    expect(img.width).toBe(120);
    expect(img.height).toBe(80);
  });

  it('throws when bytes are not a recognised image format', () => {
    expect(() => loadImage(new Uint8Array([1, 2, 3, 4, 5]))).toThrowError(/could not determine image format/);
  });

  it('honours explicit format override', () => {
    const img = loadImage(new Uint8Array([0, 0, 0]), { format: 'svg', width: 10, height: 10 });
    expect(img.format).toBe('svg');
    expect(img.width).toBe(10);
  });
});

describe('Picture drawing-xml round-trip', () => {
  it('preserves rId + name + descr through parse/serialize', () => {
    const drawing = makeDrawing([
      makePictureDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'C5' }), {
        rId: 'rId4',
        name: 'Logo',
        descr: 'Company logo',
      }),
    ]);
    const back = parseDrawingXml(drawingToBytes(drawing));
    expect(back.items.length).toBe(1);
    const item = back.items[0];
    if (!item || item.content.kind !== 'picture') throw new Error('expected picture');
    expect(item.content.picture.rId).toBe('rId4');
    expect(item.content.picture.name).toBe('Logo');
    expect(item.content.picture.descr).toBe('Company logo');
  });

  it('preserves picture across all three anchor kinds', () => {
    const items = [
      makePictureDrawingItem(makeAbsoluteAnchor({ x: 100, y: 200, cx: 300, cy: 400 }), { rId: 'rId1', name: 'abs' }),
      makePictureDrawingItem(makeOneCellAnchor({ from: 'B2', widthPx: 100, heightPx: 50 }), {
        rId: 'rId2',
        name: 'one',
      }),
      makePictureDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'C3' }), { rId: 'rId3', name: 'two' }),
    ];
    const back = parseDrawingXml(drawingToBytes(makeDrawing(items)));
    expect(back.items.map((i) => i.anchor.kind)).toEqual(['absolute', 'oneCell', 'twoCell']);
    expect(back.items.map((i) => (i.content.kind === 'picture' ? i.content.picture.name : null))).toEqual([
      'abs',
      'one',
      'two',
    ]);
  });
});

describe('Picture workbook round-trip', () => {
  it('embeds image bytes, emits manifest Default + drawing-rels image rel, restores on load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const png = loadImage(makePngBytes(640, 480));
    ws.drawing = makeDrawing([
      makePictureDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'F10' }), {
        image: png,
        name: 'Logo',
        descr: 'A 640x480 placeholder',
      }),
    ]);

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);

    // Image was placed under xl/media/image1.png.
    expect(entries['xl/media/image1.png']).toBeDefined();
    expect(entries['xl/media/image1.png']?.length).toBe(33);

    // Manifest carries a Default entry for png.
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('Extension="png"');
    expect(ct).toContain('image/png');

    // Drawing rels link to the media path with the IMAGE_REL type.
    const drelsXml = new TextDecoder().decode(entries['xl/drawings/_rels/drawing1.xml.rels']);
    expect(drelsXml).toContain('../media/image1.png');
    expect(drelsXml).toContain('relationships/image');

    // Round-trip: load the workbook back and verify the image bytes survive.
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const item = wb2.sheets[0]?.sheet.drawing?.items[0];
    if (!item || item.content.kind !== 'picture') throw new Error('expected picture');
    const back = item.content.picture.image as XlsxImage;
    expect(back.format).toBe('png');
    expect(back.width).toBe(640);
    expect(back.height).toBe(480);
    expect(back.bytes).toEqual(png.bytes);
    expect(item.content.picture.name).toBe('Logo');
    expect(item.content.picture.descr).toBe('A 640x480 placeholder');
  });

  it('handles multiple images of mixed format with workbook-global imageN ids', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    ws.drawing = makeDrawing([
      makePictureDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B2' }), {
        image: loadImage(makePngBytes(10, 10)),
      }),
      makePictureDrawingItem(makeTwoCellAnchor({ from: 'C1', to: 'D2' }), {
        image: loadImage(makeGifBytes(20, 20)),
      }),
    ]);
    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    expect(entries['xl/media/image1.png']).toBeDefined();
    expect(entries['xl/media/image2.gif']).toBeDefined();
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('Extension="png"');
    expect(ct).toContain('Extension="gif"');
    expect(ct).toContain('image/gif');
  });
});

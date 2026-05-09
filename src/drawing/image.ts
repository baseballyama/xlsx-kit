// Image embedding.
//
// `XlsxImage` is the workbook-level handle for any image referenced from a
// worksheet drawing (or chart `<a:blipFill>`). Bytes are kept verbatim so
// re-saving never re-encodes; format / width / height are detected from the
// file header so callers don't have to specify them.

import { OpenXmlIoError } from '../utils/exceptions';

export type XlsxImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tiff' | 'svg' | 'emf' | 'wmf';

/**
 * Map from image format to the Content-Types Default extension that Excel uses
 * in `[Content_Types].xml`. SVG / EMF / WMF use the same extension as the
 * format string; PNG/JPEG/GIF/BMP do too. WebP and TIFF use their canonical
 * short names.
 */
export const IMAGE_FORMAT_EXTENSION: Readonly<Record<XlsxImageFormat, string>> = {
  png: 'png',
  jpeg: 'jpeg',
  gif: 'gif',
  bmp: 'bmp',
  webp: 'webp',
  tiff: 'tiff',
  svg: 'svg',
  emf: 'emf',
  wmf: 'wmf',
};

/** Map from image format to its IANA `image/*` MIME type. */
export const IMAGE_FORMAT_MIME: Readonly<Record<XlsxImageFormat, string>> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export interface XlsxImage {
  bytes: Uint8Array;
  format: XlsxImageFormat;
  /** Pixel width. Zero when dimensions can't be determined for the format. */
  width: number;
  /** Pixel height. Zero when dimensions can't be determined for the format. */
  height: number;
  /** ZIP archive path, e.g. `xl/media/image3.png`. Set by the writer. */
  path?: string;
  /** rels-resolved id used by `<a:blip r:embed="...">`. Set by the loader. */
  rId?: string;
}

// ---- Magic-byte format detection ------------------------------------------

const startsWith = (bytes: Uint8Array, sig: ReadonlyArray<number>, offset = 0): boolean => {
  if (bytes.length < offset + sig.length) return false;
  for (let i = 0; i < sig.length; i++) {
    if (bytes[offset + i] !== sig[i]) return false;
  }
  return true;
};

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const JPEG_SIG = [0xff, 0xd8, 0xff];
const GIF87 = [0x47, 0x49, 0x46, 0x38, 0x37, 0x61];
const GIF89 = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61];
const BMP_SIG = [0x42, 0x4d];
const RIFF_SIG = [0x52, 0x49, 0x46, 0x46];
const WEBP_TAG = [0x57, 0x45, 0x42, 0x50];
const TIFF_LE = [0x49, 0x49, 0x2a, 0x00];
const TIFF_BE = [0x4d, 0x4d, 0x00, 0x2a];
const EMF_SIG = [0x01, 0x00, 0x00, 0x00];
const EMF_TAG = [0x20, 0x45, 0x4d, 0x46]; // ' EMF' at offset 40
const WMF_SIG_PLACEABLE = [0xd7, 0xcd, 0xc6, 0x9a];
const WMF_SIG_BARE = [0x01, 0x00, 0x09, 0x00];

const isSvg = (bytes: Uint8Array): boolean => {
  // Inspect the first ~512 bytes for an `<svg` tag (with optional XML / DOCTYPE
  // preamble).
  const head = new TextDecoder('utf-8', { fatal: false }).decode(bytes.subarray(0, Math.min(bytes.length, 512)));
  return /<svg\b/i.test(head);
};

export function detectImageFormat(bytes: Uint8Array): XlsxImageFormat | undefined {
  if (startsWith(bytes, PNG_SIG)) return 'png';
  if (startsWith(bytes, JPEG_SIG)) return 'jpeg';
  if (startsWith(bytes, GIF87) || startsWith(bytes, GIF89)) return 'gif';
  if (startsWith(bytes, BMP_SIG)) return 'bmp';
  if (startsWith(bytes, RIFF_SIG) && startsWith(bytes, WEBP_TAG, 8)) return 'webp';
  if (startsWith(bytes, TIFF_LE) || startsWith(bytes, TIFF_BE)) return 'tiff';
  if (startsWith(bytes, EMF_SIG) && startsWith(bytes, EMF_TAG, 40)) return 'emf';
  if (startsWith(bytes, WMF_SIG_PLACEABLE) || startsWith(bytes, WMF_SIG_BARE)) return 'wmf';
  if (isSvg(bytes)) return 'svg';
  return undefined;
}

// ---- Per-format dimension extraction --------------------------------------

const readU32BE = (bytes: Uint8Array, offset: number): number =>
  ((bytes[offset] ?? 0) << 24) |
  ((bytes[offset + 1] ?? 0) << 16) |
  ((bytes[offset + 2] ?? 0) << 8) |
  (bytes[offset + 3] ?? 0);

const readU16LE = (bytes: Uint8Array, offset: number): number => (bytes[offset] ?? 0) | ((bytes[offset + 1] ?? 0) << 8);

const readU16BE = (bytes: Uint8Array, offset: number): number => ((bytes[offset] ?? 0) << 8) | (bytes[offset + 1] ?? 0);

interface Dimensions {
  width: number;
  height: number;
}

const pngDimensions = (bytes: Uint8Array): Dimensions | undefined => {
  // PNG: 8-byte signature, 4-byte chunk-length, 4-byte type ("IHDR"), then
  // 4-byte width, 4-byte height (both big-endian).
  if (bytes.length < 24) return undefined;
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) {
    return undefined;
  }
  return { width: readU32BE(bytes, 16), height: readU32BE(bytes, 20) };
};

const gifDimensions = (bytes: Uint8Array): Dimensions | undefined => {
  if (bytes.length < 10) return undefined;
  return { width: readU16LE(bytes, 6), height: readU16LE(bytes, 8) };
};

const bmpDimensions = (bytes: Uint8Array): Dimensions | undefined => {
  if (bytes.length < 26) return undefined;
  // BITMAPINFOHEADER: width (LE int32) at offset 18, height at offset 22.
  // Height can be negative for top-down bitmaps; we report absolute pixels.
  const w = (bytes[18] ?? 0) | ((bytes[19] ?? 0) << 8) | ((bytes[20] ?? 0) << 16) | ((bytes[21] ?? 0) << 24);
  const hRaw = (bytes[22] ?? 0) | ((bytes[23] ?? 0) << 8) | ((bytes[24] ?? 0) << 16) | ((bytes[25] ?? 0) << 24);
  // Treat width as signed; for height take absolute value (top-down bitmaps
  // store negative).
  const wSigned = (w | 0) >>> 0 > 0x7fffffff ? (w | 0) - 0x100000000 : w;
  const hSigned = (hRaw | 0) >>> 0 > 0x7fffffff ? (hRaw | 0) - 0x100000000 : hRaw;
  return { width: wSigned, height: Math.abs(hSigned) };
};

const jpegDimensions = (bytes: Uint8Array): Dimensions | undefined => {
  // Walk the marker stream until an SOFn (FFC0..FFCF, excluding FFC4/C8/CC).
  let i = 2; // skip SOI (FFD8)
  while (i + 9 < bytes.length) {
    if (bytes[i] !== 0xff) return undefined;
    let marker = bytes[i + 1] ?? 0;
    while (marker === 0xff) {
      // Padding bytes. Advance to the next non-FF.
      i += 1;
      marker = bytes[i + 1] ?? 0;
    }
    i += 2;
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    const len = readU16BE(bytes, i);
    if (isSof) {
      // SOF segment: 2 bytes length, 1 byte precision, 2 bytes height, 2 bytes
      // width.
      if (i + 7 > bytes.length) return undefined;
      const height = readU16BE(bytes, i + 3);
      const width = readU16BE(bytes, i + 5);
      return { width, height };
    }
    i += len;
  }
  return undefined;
};

const webpDimensions = (bytes: Uint8Array): Dimensions | undefined => {
  // RIFF header is 12 bytes; chunk type is at offset 12.
  if (bytes.length < 30) return undefined;
  // VP8 (lossy): "VP8 " chunk; width / height at offset 26 / 28 (LE 14-bit each
  // + 2 unused bits).
  if (
    bytes[12] === 0x56 && // V
    bytes[13] === 0x50 && // P
    bytes[14] === 0x38 && // 8
    bytes[15] === 0x20 // space
  ) {
    const w = readU16LE(bytes, 26) & 0x3fff;
    const h = readU16LE(bytes, 28) & 0x3fff;
    return { width: w, height: h };
  }
  // VP8L (lossless): 14-bit width-1 / height-1 packed little-endian starting at
  // offset 21.
  if (
    bytes[12] === 0x56 &&
    bytes[13] === 0x50 &&
    bytes[14] === 0x38 &&
    bytes[15] === 0x4c // L
  ) {
    if (bytes.length < 25) return undefined;
    const b0 = bytes[21] ?? 0;
    const b1 = bytes[22] ?? 0;
    const b2 = bytes[23] ?? 0;
    const b3 = bytes[24] ?? 0;
    const w = (b0 | (b1 << 8)) & 0x3fff;
    const h = ((b1 >> 6) | (b2 << 2) | ((b3 & 0x0f) << 10)) & 0x3fff;
    return { width: w + 1, height: h + 1 };
  }
  // VP8X (extended): width-1 (24 bits) / height-1 (24 bits) at offset 24 / 27.
  if (
    bytes[12] === 0x56 &&
    bytes[13] === 0x50 &&
    bytes[14] === 0x38 &&
    bytes[15] === 0x58 // X
  ) {
    if (bytes.length < 30) return undefined;
    const w = (bytes[24] ?? 0) | ((bytes[25] ?? 0) << 8) | ((bytes[26] ?? 0) << 16);
    const h = (bytes[27] ?? 0) | ((bytes[28] ?? 0) << 8) | ((bytes[29] ?? 0) << 16);
    return { width: w + 1, height: h + 1 };
  }
  return undefined;
};

export function detectImageDimensions(bytes: Uint8Array, format: XlsxImageFormat): Dimensions {
  let dims: Dimensions | undefined;
  switch (format) {
    case 'png':
      dims = pngDimensions(bytes);
      break;
    case 'jpeg':
      dims = jpegDimensions(bytes);
      break;
    case 'gif':
      dims = gifDimensions(bytes);
      break;
    case 'bmp':
      dims = bmpDimensions(bytes);
      break;
    case 'webp':
      dims = webpDimensions(bytes);
      break;
    default:
      // tiff / svg / emf / wmf — readers can fall back to 0 and Excel will
      // assign defaults.
      dims = undefined;
  }
  return dims ?? { width: 0, height: 0 };
}

// ---- Public factory --------------------------------------------------------

/**
 * Build an `XlsxImage` from raw bytes. Detects the format and dimensions
 * automatically. Caller may pass an explicit `format` to override detection
 * (e.g. when bytes were read from a `data:` URL with a known MIME type) or
 * `width`/`height` to skip the parser.
 */
export function loadImage(
  bytes: Uint8Array,
  opts: { format?: XlsxImageFormat; width?: number; height?: number } = {},
): XlsxImage {
  const format = opts.format ?? detectImageFormat(bytes);
  if (!format) {
    throw new OpenXmlIoError('loadImage: could not determine image format from bytes');
  }
  const dims =
    opts.width !== undefined && opts.height !== undefined
      ? { width: opts.width, height: opts.height }
      : detectImageDimensions(bytes, format);
  return { bytes, format, width: dims.width, height: dims.height };
}

// DrawingML / SpreadsheetDrawing length unit conversions. EMUs (English
// Metric Units) are OOXML's universal length type — 1 inch = 914400
// EMU, 1 cm = 360000 EMU, 1 px @ 96 DPI = 9525 EMU.
//
// Mirrors openpyxl/openpyxl/utils/units.py. Per
// docs/plan/03-foundations.md §7.3 + docs/plan/01-architecture.md §7.4
// these helpers are on the hot path for the drawing / chart writers.

export const EMU_PER_INCH = 914_400;
export const EMU_PER_CM = 360_000;
export const EMU_PER_PIXEL = 9_525;
export const EMU_PER_POINT = 12_700;

/** DPI assumed by Excel when converting between pixels and other units. */
export const DEFAULT_PIXEL_DPI = 96;

// ---- pixel <-> EMU ---------------------------------------------------------

export function emuFromPx(px: number): number {
  return Math.round(px * EMU_PER_PIXEL);
}
export function pxFromEmu(emu: number): number {
  return emu / EMU_PER_PIXEL;
}

// ---- centimetre <-> EMU ----------------------------------------------------

export function emuFromCm(cm: number): number {
  return Math.round(cm * EMU_PER_CM);
}
export function cmFromEmu(emu: number): number {
  return emu / EMU_PER_CM;
}

// ---- inch <-> EMU ----------------------------------------------------------

export function emuFromInch(inch: number): number {
  return Math.round(inch * EMU_PER_INCH);
}
export function inchFromEmu(emu: number): number {
  return emu / EMU_PER_INCH;
}

// ---- point <-> EMU ---------------------------------------------------------

export function emuFromPoint(pt: number): number {
  return Math.round(pt * EMU_PER_POINT);
}
export function pointFromEmu(emu: number): number {
  return emu / EMU_PER_POINT;
}

// ---- pixel <-> point (DPI) ------------------------------------------------

export function pointToPixel(pt: number, dpi = DEFAULT_PIXEL_DPI): number {
  return (pt * dpi) / 72;
}
export function pixelToPoint(px: number, dpi = DEFAULT_PIXEL_DPI): number {
  return (px * 72) / dpi;
}

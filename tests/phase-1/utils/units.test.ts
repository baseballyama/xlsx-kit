import { describe, expect, it } from 'vitest';
import {
  cmFromEmu,
  EMU_PER_CM,
  EMU_PER_INCH,
  EMU_PER_PIXEL,
  EMU_PER_POINT,
  emuFromCm,
  emuFromInch,
  emuFromPoint,
  emuFromPx,
  inchFromEmu,
  pixelToPoint,
  pointFromEmu,
  pointToPixel,
  pxFromEmu,
} from '../../../src/utils/units';

describe('EMU constants match the OOXML / openpyxl reference', () => {
  it('1 inch = 914400 EMU', () => expect(EMU_PER_INCH).toBe(914_400));
  it('1 cm = 360000 EMU', () => expect(EMU_PER_CM).toBe(360_000));
  it('1 px @ 96 DPI = 9525 EMU', () => expect(EMU_PER_PIXEL).toBe(9_525));
  it('1 pt = 12700 EMU', () => expect(EMU_PER_POINT).toBe(12_700));
});

describe('EMU conversions round-trip', () => {
  it('emuFromPx / pxFromEmu', () => {
    expect(emuFromPx(96)).toBe(914_400);
    expect(pxFromEmu(914_400)).toBe(96);
  });

  it('emuFromCm / cmFromEmu', () => {
    expect(emuFromCm(2.54)).toBeCloseTo(914_400, 6);
    expect(cmFromEmu(360_000)).toBe(1);
  });

  it('emuFromInch / inchFromEmu', () => {
    expect(emuFromInch(1)).toBe(914_400);
    expect(inchFromEmu(914_400)).toBe(1);
  });

  it('emuFromPoint / pointFromEmu', () => {
    expect(emuFromPoint(72)).toBe(914_400);
    expect(pointFromEmu(914_400)).toBe(72);
  });
});

describe('point <-> pixel via DPI', () => {
  it('72 pt = 96 px @ default 96 DPI', () => {
    expect(pointToPixel(72)).toBe(96);
    expect(pixelToPoint(96)).toBe(72);
  });

  it('honours a custom DPI', () => {
    expect(pointToPixel(72, 144)).toBe(144);
    expect(pixelToPoint(144, 144)).toBe(72);
  });
});

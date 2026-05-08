import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../../src/schema/serialize';
import {
  BLACK,
  BLUE,
  COLOR_INDEX,
  makeColor,
  normaliseRgb,
  resolveIndexedColor,
  rgbColor,
  WHITE,
} from '../../../src/xlsx/styles/colors';
import { ColorSchema } from '../../../src/xlsx/styles/colors.schema';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('Color value object', () => {
  it('makeColor returns a frozen object', () => {
    const c = makeColor({ rgb: 'FF0000' });
    expect(Object.isFrozen(c)).toBe(true);
  });

  it('omits unset fields rather than carrying undefined', () => {
    const c = makeColor({ theme: 1 });
    expect(Object.keys(c).sort()).toEqual(['theme']);
  });

  it('normalises 6-digit rgb to 8-digit uppercase aRGB', () => {
    expect(makeColor({ rgb: 'ff0000' }).rgb).toBe('00FF0000');
    expect(makeColor({ rgb: 'FF0000' }).rgb).toBe('00FF0000');
    expect(makeColor({ rgb: 'FFFF0000' }).rgb).toBe('FFFF0000');
  });

  it('rejects malformed rgb', () => {
    expect(() => makeColor({ rgb: 'abc' })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ rgb: 'GG0000' })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ rgb: '' })).toThrowError(OpenXmlSchemaError);
  });

  it('rejects out-of-range indexed / theme / tint', () => {
    expect(() => makeColor({ indexed: -1 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ indexed: 66 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ theme: -1 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ tint: 1.5 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeColor({ tint: -1.1 })).toThrowError(OpenXmlSchemaError);
  });

  it('accepts the boundary tint values -1 and +1', () => {
    expect(makeColor({ tint: -1 }).tint).toBe(-1);
    expect(makeColor({ tint: 1 }).tint).toBe(1);
  });

  it('rgbColor is a thin makeColor shortcut', () => {
    expect(rgbColor('FF0000')).toEqual(makeColor({ rgb: 'FF0000' }));
  });

  it('normaliseRgb is exposed and idempotent', () => {
    expect(normaliseRgb('FF0000')).toBe('00FF0000');
    expect(normaliseRgb('00FF0000')).toBe('00FF0000');
  });
});

describe('COLOR_INDEX palette', () => {
  it('has exactly 64 entries (legacy openpyxl table)', () => {
    expect(COLOR_INDEX.length).toBe(64);
    expect(Object.isFrozen(COLOR_INDEX)).toBe(true);
  });

  it('matches the openpyxl named exports', () => {
    expect(BLACK).toBe('00000000');
    expect(WHITE).toBe('00FFFFFF');
    expect(BLUE).toBe('000000FF');
  });

  it('resolveIndexedColor returns the entry or undefined', () => {
    expect(resolveIndexedColor(0)).toBe('00000000');
    expect(resolveIndexedColor(63)).toBe('00333333');
    expect(resolveIndexedColor(64)).toBeUndefined(); // system fg
    expect(resolveIndexedColor(99)).toBeUndefined();
  });
});

describe('Color schema round-trip', () => {
  it('round-trips an rgb-only color via XML', () => {
    const c = makeColor({ rgb: 'FF112233' });
    const back = fromTree(parseXml(serializeXml(toTree(c, ColorSchema))), ColorSchema);
    expect(back).toEqual(c);
  });

  it('round-trips a theme + tint color', () => {
    const c = makeColor({ theme: 1, tint: -0.25 });
    const back = fromTree(parseXml(serializeXml(toTree(c, ColorSchema))), ColorSchema);
    expect(back).toEqual(c);
  });

  it('round-trips an auto color', () => {
    const c = makeColor({ auto: true });
    const back = fromTree(parseXml(serializeXml(toTree(c, ColorSchema))), ColorSchema);
    expect(back).toEqual(c);
  });

  it('round-trips an indexed color', () => {
    const c = makeColor({ indexed: 64 });
    const back = fromTree(parseXml(serializeXml(toTree(c, ColorSchema))), ColorSchema);
    expect(back).toEqual(c);
  });
});

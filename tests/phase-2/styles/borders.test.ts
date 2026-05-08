import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../../src/schema/serialize';
import { DEFAULT_BORDER, EMPTY_SIDE, makeBorder, makeSide, SIDE_STYLES } from '../../../src/styles/borders';
import { BorderSchema, SideSchema } from '../../../src/styles/borders.schema';
import { makeColor } from '../../../src/styles/colors';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('Side', () => {
  it('makeSide returns a frozen object', () => {
    const s = makeSide({ style: 'thin' });
    expect(Object.isFrozen(s)).toBe(true);
  });

  it('accepts every legal style', () => {
    for (const style of SIDE_STYLES) {
      expect(makeSide({ style }).style).toBe(style);
    }
  });

  it('rejects an unknown style', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => makeSide({ style: 'zigzag' as any })).toThrowError(OpenXmlSchemaError);
  });

  it('freezes nested color even when caller passes a plain object', () => {
    const s = makeSide({ style: 'thin', color: { rgb: 'FF000000' } });
    expect(s.color && Object.isFrozen(s.color)).toBe(true);
    expect(s.color?.rgb).toBe('FF000000');
  });

  it('round-trips via the schema layer', () => {
    const s = makeSide({ style: 'medium', color: makeColor({ rgb: 'FFAABBCC' }) });
    const back = fromTree(parseXml(serializeXml(toTree(s, SideSchema))), SideSchema);
    expect(back).toEqual(s);
  });
});

describe('Border', () => {
  it('makeBorder returns a frozen object and freezes nested sides', () => {
    const b = makeBorder({ left: { style: 'thin' }, outline: true });
    expect(Object.isFrozen(b)).toBe(true);
    expect(b.left && Object.isFrozen(b.left)).toBe(true);
  });

  it('DEFAULT_BORDER is the empty Border value', () => {
    expect(DEFAULT_BORDER).toEqual({});
    expect(Object.isFrozen(DEFAULT_BORDER)).toBe(true);
  });

  it('EMPTY_SIDE is the empty Side value', () => {
    expect(EMPTY_SIDE).toEqual({});
    expect(Object.isFrozen(EMPTY_SIDE)).toBe(true);
  });

  it('serializes outline + nested sides into expected XML', () => {
    const border = makeBorder({
      outline: true,
      left: makeSide({ style: 'thin', color: makeColor({ rgb: 'FF112233' }) }),
      bottom: makeSide({ style: 'medium' }),
    });
    const xml = new TextDecoder().decode(serializeXml(toTree(border, BorderSchema)));
    expect(xml).toContain('<border');
    expect(xml).toContain('outline="1"');
    expect(xml).toContain('<left style="thin">');
    expect(xml).toContain('rgb="FF112233"');
    expect(xml).toContain('<bottom style="medium"/>');
  });

  it('round-trips a fully populated Border via XML', () => {
    const border = makeBorder({
      diagonalUp: true,
      diagonalDown: false,
      outline: true,
      left: makeSide({ style: 'thin', color: makeColor({ rgb: 'FFAABBCC' }) }),
      right: makeSide({ style: 'medium' }),
      top: makeSide({ style: 'thick' }),
      bottom: makeSide({ style: 'double' }),
      diagonal: makeSide({ style: 'dotted' }),
    });
    const back = fromTree(parseXml(serializeXml(toTree(border, BorderSchema))), BorderSchema);
    expect(back).toEqual(border);
  });

  it('parses openpyxl-style XML loosely', () => {
    const xml = `<?xml version="1.0"?><border xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" diagonalUp="true" outline="false"><left style="thin"><color rgb="FF000000"/></left></border>`;
    const b = fromTree(parseXml(xml), BorderSchema);
    expect(b.diagonalUp).toBe(true);
    expect(b.outline).toBe(false);
    expect(b.left?.style).toBe('thin');
    expect(b.left?.color?.rgb).toBe('FF000000');
  });
});

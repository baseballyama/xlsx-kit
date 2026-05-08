import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { makeFont } from '../../src/styles/fonts';
import { addCellXf, addFont } from '../../src/styles/stylesheet';
import { parseStylesheetXml } from '../../src/styles/stylesheet-reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const wrap = (sections: string): string =>
  `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${sections}</styleSheet>`;

describe('parseStylesheetXml — section ordering', () => {
  it('preserves slot indices across all pools', () => {
    const xml = wrap(
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="14"/><name val="Arial"/><b/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="1" fillId="2" borderId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/></cellXfs>',
    );
    const ss = parseStylesheetXml(xml);
    expect(ss.fonts.length).toBe(2);
    expect(ss.fonts[0]?.size).toBe(11);
    expect(ss.fonts[0]?.name).toBe('Calibri');
    expect(ss.fonts[1]?.size).toBe(14);
    expect(ss.fonts[1]?.bold).toBe(true);
    expect(ss.fills.length).toBe(3);
    expect(ss.fills[2]?.kind).toBe('pattern');
    expect(ss.borders.length).toBe(1);
    expect(ss.cellXfs.length).toBe(2);
    expect(ss.cellXfs[1]?.numFmtId).toBe(14);
    expect(ss.cellXfs[1]?.fontId).toBe(1);
    expect(ss.cellXfs[1]?.fillId).toBe(2);
    expect(ss.cellXfs[1]?.applyFont).toBe(true);
    expect(ss.cellXfs[1]?.applyNumberFormat).toBe(true);
  });

  it('parses cellStyleXfs separately from cellXfs', () => {
    const xml = wrap(
      '<fonts count="1"><font><sz val="11"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>',
    );
    const ss = parseStylesheetXml(xml);
    expect(ss.cellStyleXfs.length).toBe(1);
    expect(ss.cellXfs.length).toBe(2);
    expect(ss.cellXfs[0]?.xfId).toBe(0);
    expect(ss.cellXfs[1]?.xfId).toBe(0);
  });

  it('parses <numFmts> into the Map<id, code>', () => {
    const xml = wrap(
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0000"/><numFmt numFmtId="165" formatCode="0.00%"/></numFmts>' +
        '<fonts count="1"><font><sz val="11"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellXfs count="1"><xf numFmtId="164" fontId="0" fillId="0" borderId="0"/></cellXfs>',
    );
    const ss = parseStylesheetXml(xml);
    expect(ss.numFmts.size).toBe(2);
    expect(ss.numFmts.get(164)).toBe('0.0000');
    expect(ss.numFmts.get(165)).toBe('0.00%');
  });

  it('parses inline <alignment> and <protection> children of <xf>', () => {
    const xml = wrap(
      '<fonts count="1"><font><sz val="11"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="top"/><protection locked="0" hidden="1"/></xf></cellXfs>',
    );
    const ss = parseStylesheetXml(xml);
    const xf = ss.cellXfs[0];
    expect(xf?.alignment?.horizontal).toBe('center');
    expect(xf?.alignment?.vertical).toBe('top');
    expect(xf?.protection?.locked).toBe(false);
    expect(xf?.protection?.hidden).toBe(true);
  });

  it('rejects a non-styleSheet root', () => {
    expect(() => parseStylesheetXml('<foo/>')).toThrowError(/expected styleSheet/);
  });
});

describe('parseStylesheetXml — index rebuild', () => {
  it('post-load addFont dedupes against pre-existing entries', () => {
    const xml = wrap(
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="14"/><name val="Arial"/><b/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    );
    const ss = parseStylesheetXml(xml);
    // Re-adding the bold Arial 14 must hit slot 1, not allocate a new slot.
    expect(addFont(ss, makeFont({ name: 'Arial', size: 14, bold: true }))).toBe(1);
    expect(ss.fonts.length).toBe(2);
  });

  it('post-load addCellXf dedupes structurally-equivalent xfs', () => {
    const xml = wrap(
      '<fonts count="1"><font><sz val="11"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>',
    );
    const ss = parseStylesheetXml(xml);
    // Same shape as cellXfs[1] should land back at id 1.
    expect(
      addCellXf(ss, {
        fontId: 0,
        fillId: 0,
        borderId: 0,
        numFmtId: 9,
        applyNumberFormat: true,
      }),
    ).toBe(1);
    expect(ss.cellXfs.length).toBe(2);
  });
});

describe('parseStylesheetXml — fixture round-trip', () => {
  it('reads openpyxl genuine/empty-with-styles.xlsx styles.xml', async () => {
    const { unzipSync } = await import('fflate');
    const xlsx = readFileSync(resolve(FIXTURES, 'empty-with-styles.xlsx'));
    const entries = unzipSync(xlsx);
    const stylesBytes = entries['xl/styles.xml'];
    if (!stylesBytes) throw new Error('expected xl/styles.xml');
    const ss = parseStylesheetXml(stylesBytes);
    expect(ss.fonts.length).toBe(1);
    expect(ss.fills.length).toBe(2);
    expect(ss.borders.length).toBe(1);
    expect(ss.cellStyleXfs.length).toBe(1);
    expect(ss.cellXfs.length).toBe(5);
    // Built-in numFmtIds 10/14/20/2 etc.
    expect(ss.cellXfs[1]?.numFmtId).toBe(10);
    expect(ss.cellXfs[2]?.numFmtId).toBe(14);
    expect(ss.cellXfs[3]?.numFmtId).toBe(20);
    expect(ss.cellXfs[4]?.numFmtId).toBe(2);
  });
});

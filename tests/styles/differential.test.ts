import { describe, expect, it } from 'vitest';
import { makeAlignment } from '../../src/styles/alignment';
import { makeColor } from '../../src/styles/colors';
import { addDxf, getDxfs, makeDifferentialStyle } from '../../src/styles/differential';
import { makePatternFill } from '../../src/styles/fills';
import { makeFont } from '../../src/styles/fonts';
import { makeNumberFormat } from '../../src/styles/numbers';
import { makeProtection } from '../../src/styles/protection';
import { makeStylesheet } from '../../src/styles/stylesheet';

describe('DifferentialStyle', () => {
  it('makeDifferentialStyle freezes the result and omits unset fields', () => {
    const d = makeDifferentialStyle({ font: makeFont({ bold: true }) });
    expect(Object.isFrozen(d)).toBe(true);
    expect(Object.keys(d)).toEqual(['font']);
  });

  it('round-trips its full surface', () => {
    const d = makeDifferentialStyle({
      font: makeFont({ italic: true, color: makeColor({ rgb: 'FF0000' }) }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: '00FF00' }) }),
      alignment: makeAlignment({ horizontal: 'center' }),
      protection: makeProtection({ locked: false }),
      numFmt: makeNumberFormat({ numFmtId: 200, formatCode: '0.0%' }),
    });
    expect(d.font?.italic).toBe(true);
    expect(d.fill?.kind).toBe('pattern');
    expect(d.alignment?.horizontal).toBe('center');
    expect(d.protection?.locked).toBe(false);
    expect(d.numFmt?.formatCode).toBe('0.0%');
  });
});

describe('addDxf — Stylesheet pool extension', () => {
  it('lazy-allocates the dxfs array on first add', () => {
    const ss = makeStylesheet();
    expect(getDxfs(ss)).toEqual([]);
    addDxf(ss, makeDifferentialStyle({ font: makeFont({ bold: true }) }));
    expect(getDxfs(ss).length).toBe(1);
  });

  it('is idempotent on structural equality', () => {
    const ss = makeStylesheet();
    const a = addDxf(ss, makeDifferentialStyle({ font: makeFont({ bold: true }) }));
    const b = addDxf(ss, makeDifferentialStyle({ font: makeFont({ bold: true }) }));
    expect(a).toBe(b);
    expect(getDxfs(ss).length).toBe(1);
  });

  it('different DXFs allocate distinct indices', () => {
    const ss = makeStylesheet();
    const bold = addDxf(ss, makeDifferentialStyle({ font: makeFont({ bold: true }) }));
    const italic = addDxf(ss, makeDifferentialStyle({ font: makeFont({ italic: true }) }));
    const filled = addDxf(
      ss,
      makeDifferentialStyle({ fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FF0000' }) }) }),
    );
    expect(new Set([bold, italic, filled]).size).toBe(3);
    expect(getDxfs(ss).length).toBe(3);
  });

  it('insertion-order independent dedup via stableStringify', () => {
    const ss = makeStylesheet();
    const a = addDxf(
      ss,
      makeDifferentialStyle({
        font: makeFont({ bold: true, italic: true }),
        alignment: makeAlignment({ horizontal: 'center', wrapText: true }),
      }),
    );
    const b = addDxf(
      ss,
      makeDifferentialStyle({
        alignment: makeAlignment({ wrapText: true, horizontal: 'center' }),
        font: makeFont({ italic: true, bold: true }),
      }),
    );
    expect(a).toBe(b);
  });
});

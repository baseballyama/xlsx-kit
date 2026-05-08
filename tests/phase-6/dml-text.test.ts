import { describe, expect, it } from 'vitest';
import { makeColor, makeSchemeColor, makeSrgbColor } from '../../src/xlsx/drawing/dml/colors';
import {
  parseTextBody,
  parseTextBodyProperties,
  serializeTextBody,
  serializeTextBodyProperties,
} from '../../src/xlsx/drawing/dml/dml-xml';
import { makeEffectList } from '../../src/xlsx/drawing/dml/effect';
import { makeNoFill, makeSolidFill } from '../../src/xlsx/drawing/dml/fill';
import { makeLine } from '../../src/xlsx/drawing/dml/line';
import {
  makeBreak,
  makeParagraph,
  makeRun,
  makeRunProperties,
  makeSimpleTextBody,
  makeTextBody,
  type ParagraphProperties,
  type RunProperties,
  type TextBody,
  type TextBodyProperties,
} from '../../src/xlsx/drawing/dml/text';
import { parseXml } from '../../src/xml/parser';
import { findChild } from '../../src/xml/tree';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NSDECL = `xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"`;

const roundTripBody = (b: TextBody): TextBody => {
  const xml = `<root ${NSDECL}>${serializeTextBody(b)}</root>`;
  const root = parseXml(xml);
  const txEl = findChild(root, `{${C_NS}}txPr`);
  if (!txEl) throw new Error('text body round-trip: <c:txPr> missing');
  return parseTextBody(txEl);
};

const roundTripBodyPr = (p: TextBodyProperties): TextBodyProperties => {
  const xml = `<root ${NSDECL}>${serializeTextBodyProperties('a:bodyPr', p)}</root>`;
  const root = parseXml(xml);
  const el = findChild(root, `{${A_NS}}bodyPr`);
  if (!el) throw new Error('bodyPr round-trip: missing element');
  return parseTextBodyProperties(el);
};

describe('TextBody simple cases', () => {
  it('preserves a single paragraph + single run', () => {
    const back = roundTripBody(makeSimpleTextBody('Hello'));
    expect(back.paragraphs.length).toBe(1);
    expect(back.paragraphs[0]?.runs.length).toBe(1);
    const run = back.paragraphs[0]?.runs[0];
    if (!run || run.kind !== 'r') throw new Error('expected run');
    expect(run.t).toBe('Hello');
  });

  it('preserves run-property attributes (sz/b/i/u/strike/cap/baseline)', () => {
    const rPr: RunProperties = {
      sz: 1100,
      b: true,
      i: true,
      u: 'sng',
      strike: 'dblStrike',
      cap: 'all',
      baseline: 30000,
      kern: 1200,
      lang: 'en-US',
      altLang: 'ja-JP',
      noProof: false,
      dirty: true,
    };
    const body = makeTextBody([makeParagraph([makeRun('Styled', rPr)])]);
    const back = roundTripBody(body);
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.rPr).toEqual(rPr);
  });

  it('preserves nested rPr fill / line / latin font', () => {
    const rPr = makeRunProperties({
      sz: 1400,
      b: true,
      fill: makeSolidFill(makeColor(makeSchemeColor('accent2'), [{ kind: 'lumMod', val: 75000 }])),
      ln: makeLine({ w: 6350, fill: makeSolidFill(makeColor(makeSrgbColor('FF0000'))) }),
      latin: { typeface: 'Calibri', pitchFamily: 34 },
      ea: { typeface: 'Yu Gothic' },
    });
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('Decorated', rPr)])]));
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.rPr?.fill).toEqual(rPr.fill);
    expect(r.rPr?.ln?.w).toBe(6350);
    expect(r.rPr?.latin).toEqual({ typeface: 'Calibri', pitchFamily: 34 });
    expect(r.rPr?.ea).toEqual({ typeface: 'Yu Gothic' });
  });

  it('preserves uLn=follow + uFill explicit', () => {
    const rPr: RunProperties = {
      uLn: 'follow',
      uFill: makeSolidFill(makeColor(makeSrgbColor('00FF00'))),
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('U', rPr)])]));
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.rPr?.uLn).toBe('follow');
    expect(r.rPr?.uFill).toEqual(makeSolidFill(makeColor(makeSrgbColor('00FF00'))));
  });

  it('preserves hyperlinkClick with rId + tooltip', () => {
    const rPr: RunProperties = {
      hlinkClick: { rId: 'rId7', tooltip: 'Open', history: false, highlightClick: true },
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('Link', rPr)])]));
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.rPr?.hlinkClick).toEqual({
      rId: 'rId7',
      tooltip: 'Open',
      history: false,
      highlightClick: true,
    });
  });
});

describe('TextRun variants', () => {
  it('preserves <a:br> with rPr', () => {
    const body = makeTextBody([
      makeParagraph([makeRun('a', { sz: 1100 }), makeBreak({ sz: 1100 }), makeRun('b', { sz: 1100 })]),
    ]);
    const back = roundTripBody(body);
    expect(back.paragraphs[0]?.runs.map((r) => r.kind)).toEqual(['r', 'br', 'r']);
    const br = back.paragraphs[0]?.runs[1];
    if (!br || br.kind !== 'br') throw new Error('expected br');
    expect(br.rPr?.sz).toBe(1100);
  });

  it('preserves <a:fld> with id + type + text', () => {
    const body: TextBody = makeTextBody([
      {
        runs: [{ kind: 'fld', id: '{1234-ABCD}', type: 'datetime1', t: '2026-05-04', rPr: { sz: 1000 } }],
      },
    ]);
    const back = roundTripBody(body);
    const fld = back.paragraphs[0]?.runs[0];
    if (!fld || fld.kind !== 'fld') throw new Error('expected fld');
    expect(fld.id).toBe('{1234-ABCD}');
    expect(fld.type).toBe('datetime1');
    expect(fld.t).toBe('2026-05-04');
    expect(fld.rPr?.sz).toBe(1000);
  });
});

describe('Paragraph properties', () => {
  it('preserves algn / marL / marR / lvl / indent / defTabSz', () => {
    const pPr: ParagraphProperties = {
      algn: 'ctr',
      marL: 100,
      marR: 200,
      lvl: 2,
      indent: 360,
      defTabSz: 914400,
      rtl: false,
      eaLnBrk: true,
      hangingPunct: true,
      fontAlgn: 'ctr',
    };
    const body = makeTextBody([makeParagraph([makeRun('a')], pPr)]);
    const back = roundTripBody(body);
    expect(back.paragraphs[0]?.pPr).toEqual(pPr);
  });

  it('preserves lnSpc/spcBef/spcAft (pct vs pts)', () => {
    const pPr: ParagraphProperties = {
      lnSpc: { kind: 'pct', val: 100000 },
      spcBef: { kind: 'pts', val: 600 },
      spcAft: { kind: 'pct', val: 0 },
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('x')], pPr)]));
    expect(back.paragraphs[0]?.pPr).toEqual(pPr);
  });

  it('preserves bullet (autoNum + buFont follow + buClr explicit)', () => {
    const pPr: ParagraphProperties = {
      bullet: { kind: 'autoNum', type: 'arabicPlain', startAt: 1 },
      buFont: 'follow',
      buClr: makeColor(makeSrgbColor('AABBCC')),
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('x')], pPr)]));
    expect(back.paragraphs[0]?.pPr?.bullet).toEqual({ kind: 'autoNum', type: 'arabicPlain', startAt: 1 });
    expect(back.paragraphs[0]?.pPr?.buFont).toBe('follow');
    expect(back.paragraphs[0]?.pPr?.buClr).toEqual(makeColor(makeSrgbColor('AABBCC')));
  });

  it('preserves bullet (char + buNone variants)', () => {
    const back1 = roundTripBody(makeTextBody([makeParagraph([makeRun('a')], { bullet: { kind: 'char', char: '•' } })]));
    expect(back1.paragraphs[0]?.pPr?.bullet).toEqual({ kind: 'char', char: '•' });
    const back2 = roundTripBody(makeTextBody([makeParagraph([makeRun('a')], { bullet: { kind: 'none' } })]));
    expect(back2.paragraphs[0]?.pPr?.bullet).toEqual({ kind: 'none' });
  });

  it('preserves tab stops + defRPr', () => {
    const pPr: ParagraphProperties = {
      tabLst: [
        { pos: 720, algn: 'l' },
        { pos: 1440, algn: 'ctr' },
        { pos: 2160, algn: 'r' },
      ],
      defRPr: { sz: 1000, fill: makeSolidFill(makeColor(makeSrgbColor('000000'))) },
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('x')], pPr)]));
    expect(back.paragraphs[0]?.pPr?.tabLst).toEqual(pPr.tabLst);
    expect(back.paragraphs[0]?.pPr?.defRPr?.sz).toBe(1000);
  });
});

describe('TextBodyProperties (bodyPr)', () => {
  it('preserves rot / wrap / lIns / vert / anchor / anchorCtr', () => {
    const bp: TextBodyProperties = {
      rot: 5400000,
      spcFirstLastPara: true,
      vertOverflow: 'clip',
      horzOverflow: 'overflow',
      vert: 'eaVert',
      wrap: 'square',
      lIns: 91440,
      tIns: 45720,
      rIns: 91440,
      bIns: 45720,
      numCol: 1,
      anchor: 'ctr',
      anchorCtr: false,
      forceAA: true,
      compatLnSpc: false,
    };
    expect(roundTripBodyPr(bp)).toEqual(bp);
  });

  it('preserves autoFit normAutofit with fontScale + lnSpcReduction', () => {
    const bp: TextBodyProperties = {
      autoFit: { kind: 'normAutofit', fontScale: 90000, lnSpcReduction: 10000 },
    };
    expect(roundTripBodyPr(bp)).toEqual(bp);
  });

  it('preserves spAutoFit and noAutofit', () => {
    expect(roundTripBodyPr({ autoFit: { kind: 'spAutoFit' } })).toEqual({ autoFit: { kind: 'spAutoFit' } });
    expect(roundTripBodyPr({ autoFit: { kind: 'noAutofit' } })).toEqual({ autoFit: { kind: 'noAutofit' } });
  });
});

describe('Multiple paragraphs + endParaRPr', () => {
  it('preserves multi-paragraph structure with endParaRPr', () => {
    const body: TextBody = makeTextBody([
      makeParagraph([makeRun('Title')], { algn: 'ctr' }, { sz: 1800, b: true }),
      makeParagraph([makeRun('Body line one'), makeBreak(), makeRun('Body line two')], { algn: 'l' }),
    ]);
    const back = roundTripBody(body);
    expect(back.paragraphs.length).toBe(2);
    expect(back.paragraphs[0]?.endParaRPr).toEqual({ sz: 1800, b: true });
    expect(back.paragraphs[1]?.runs.map((r) => r.kind)).toEqual(['r', 'br', 'r']);
  });
});

describe('lstStyle round-trip', () => {
  it('preserves defPPr + lvl1pPr', () => {
    const body: TextBody = makeTextBody(
      [makeParagraph([makeRun('x')])],
      {},
      {
        defPPr: { algn: 'l', defRPr: { sz: 1000 } },
        lvl1pPr: { algn: 'ctr', defRPr: { sz: 1200, b: true } },
      },
    );
    const back = roundTripBody(body);
    expect(back.lstStyle?.defPPr?.algn).toBe('l');
    expect(back.lstStyle?.defPPr?.defRPr?.sz).toBe(1000);
    expect(back.lstStyle?.lvl1pPr?.defRPr?.b).toBe(true);
  });
});

describe('Unused-but-present escape edge cases', () => {
  it('escapes < > & in run text', () => {
    const back = roundTripBody(makeSimpleTextBody('1 < 2 & 3 > 0'));
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.t).toBe('1 < 2 & 3 > 0');
  });

  it('round-trips effects on rPr', () => {
    const rPr: RunProperties = {
      sz: 1100,
      effects: { kind: 'lst', list: makeEffectList([{ kind: 'softEdge', rad: 12700 }]) },
      fill: makeNoFill(),
    };
    const back = roundTripBody(makeTextBody([makeParagraph([makeRun('x', rPr)])]));
    const r = back.paragraphs[0]?.runs[0];
    if (!r || r.kind !== 'r') throw new Error('expected run');
    expect(r.rPr?.effects).toEqual(rPr.effects);
    expect(r.rPr?.fill).toEqual(rPr.fill);
  });
});

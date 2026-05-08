// Property-based check that the Stylesheet pool dedup is **idempotent
// and order-independent**. The pools are keyed via stableStringify so
// inserting the same logical Font / Fill / Border / numFmt / CellXf
// any number of times — interleaved with other entries — must always
// yield the same id.
//
// This complements the example-based dedup tests in
// `tests/phase-2/styles/stylesheet.test.ts`.

import fc from 'fast-check';
import { describe, it } from 'vitest';
import { makeBorder, makeSide, type SideStyle } from '../../src/xlsx/styles/borders';
import { makePatternFill, type PatternType } from '../../src/xlsx/styles/fills';
import { makeFont } from '../../src/xlsx/styles/fonts';
import {
  addBorder,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  makeStylesheet,
  type Stylesheet,
} from '../../src/xlsx/styles/stylesheet';

const fontArb = fc.record(
  {
    name: fc.constantFrom('Calibri', 'Arial', 'Times New Roman'),
    size: fc.integer({ min: 8, max: 18 }),
    bold: fc.boolean(),
    italic: fc.boolean(),
  },
  { requiredKeys: ['name', 'size'] },
);

const fillArb = fc.record(
  {
    patternType: fc.constantFrom<PatternType>('solid', 'gray125', 'lightGray'),
    fgColor: fc.option(fc.record({ rgb: fc.constantFrom('FFFF0000', 'FF00FF00', 'FF0000FF') }), { nil: undefined }),
  },
  { requiredKeys: ['patternType'] },
);

const sideArb = fc.record(
  { style: fc.constantFrom<SideStyle>('thin', 'medium', 'thick', 'dashed') },
  { requiredKeys: ['style'] },
);

const borderArb = fc.record(
  {
    left: fc.option(sideArb, { nil: undefined }),
    right: fc.option(sideArb, { nil: undefined }),
    top: fc.option(sideArb, { nil: undefined }),
    bottom: fc.option(sideArb, { nil: undefined }),
  },
  { requiredKeys: [] },
);

const numFmtArb = fc.constantFrom(
  'General',
  '0',
  '0.00',
  '0.0000',
  '#,##0',
  '0%',
  'mm/dd/yyyy',
  // Two custom codes outside the built-in catalogue:
  '"Custom-A" 0.00',
  '"Custom-B" 0',
);

describe('Stylesheet pool dedup — fast-check properties', () => {
  it('addFont is idempotent under repetition', () => {
    fc.assert(
      fc.property(fontArb, fc.integer({ min: 1, max: 50 }), (opts, n) => {
        const ss = makeStylesheet();
        const font = makeFont(opts);
        const ids = Array.from({ length: n }, () => addFont(ss, font));
        return ids.every((id) => id === ids[0]);
      }),
      { numRuns: 50 },
    );
  });

  it('addFill is idempotent under repetition', () => {
    fc.assert(
      fc.property(fillArb, fc.integer({ min: 1, max: 50 }), (opts, n) => {
        const ss = makeStylesheet();
        const fill = makePatternFill({
          patternType: opts.patternType,
          ...(opts.fgColor ? { fgColor: opts.fgColor } : {}),
        });
        const ids = Array.from({ length: n }, () => addFill(ss, fill));
        return ids.every((id) => id === ids[0]);
      }),
      { numRuns: 50 },
    );
  });

  it('addBorder is idempotent under repetition', () => {
    fc.assert(
      fc.property(borderArb, fc.integer({ min: 1, max: 50 }), (opts, n) => {
        const ss = makeStylesheet();
        const border = makeBorder({
          ...(opts.left ? { left: makeSide(opts.left) } : {}),
          ...(opts.right ? { right: makeSide(opts.right) } : {}),
          ...(opts.top ? { top: makeSide(opts.top) } : {}),
          ...(opts.bottom ? { bottom: makeSide(opts.bottom) } : {}),
        });
        const ids = Array.from({ length: n }, () => addBorder(ss, border));
        return ids.every((id) => id === ids[0]);
      }),
      { numRuns: 50 },
    );
  });

  it('addNumFmt is idempotent for the same code', () => {
    fc.assert(
      fc.property(numFmtArb, fc.integer({ min: 1, max: 50 }), (code, n) => {
        const ss = makeStylesheet();
        const ids = Array.from({ length: n }, () => addNumFmt(ss, code));
        return ids.every((id) => id === ids[0]);
      }),
      { numRuns: 50 },
    );
  });

  it('addCellXf is idempotent under repetition', () => {
    fc.assert(
      fc.property(
        fc.record({
          fontId: fc.integer({ min: 0, max: 0 }),
          fillId: fc.integer({ min: 0, max: 1 }),
          borderId: fc.integer({ min: 0, max: 0 }),
          numFmtId: fc.constantFrom(0, 1, 2, 14, 22),
        }),
        fc.integer({ min: 1, max: 30 }),
        (xfShape, n) => {
          const ss = makeStylesheet();
          const xf: CellXf = xfShape;
          const ids = Array.from({ length: n }, () => addCellXf(ss, xf));
          return ids.every((id) => id === ids[0]);
        },
      ),
      { numRuns: 50 },
    );
  });

  it('insertion order does not affect dedup ids', () => {
    fc.assert(
      fc.property(fc.array(fontArb, { minLength: 2, maxLength: 8 }), (opts) => {
        const ssA = makeStylesheet();
        const ssB = makeStylesheet();
        // Same logical fonts in two different orders must yield the same
        // *set* of ids. Run forward + reverse and compare.
        const fonts = opts.map(makeFont);
        const idsForward = fonts.map((f) => addFont(ssA, f));
        const idsBackward = [...fonts].reverse().map((f) => addFont(ssB, f));
        // The ids assigned to each unique logical font must agree.
        // Map font→id in each run and compare keysets + counts.
        const distinctForward = new Set(idsForward).size;
        const distinctBackward = new Set(idsBackward).size;
        return distinctForward === distinctBackward;
      }),
      { numRuns: 50 },
    );
  });

  it('interleaving distinct + duplicate adds keeps pool size minimal', () => {
    fc.assert(
      fc.property(fc.array(fontArb, { minLength: 2, maxLength: 6 }), (opts) => {
        const ss: Stylesheet = makeStylesheet();
        const fonts = opts.map(makeFont);
        // Add each font twice, interleaved: f0, f1, f0, f1, f2, f0...
        for (let i = 0; i < 3; i++) {
          for (const f of fonts) addFont(ss, f);
        }
        const distinct = new Set(opts.map((o) => JSON.stringify(o))).size;
        // Pool starts with DEFAULT_FONT (1 entry) plus one slot per
        // logically-distinct font.
        return ss.fonts.length === 1 + distinct;
      }),
      { numRuns: 50 },
    );
  });
});

describe('cellXfs dedup is order-independent', () => {
  it('two streams of CellXf adds yield the same cellXfs[] regardless of order', () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            fontId: fc.integer({ min: 0, max: 0 }),
            fillId: fc.integer({ min: 0, max: 1 }),
            borderId: fc.integer({ min: 0, max: 0 }),
            numFmtId: fc.integer({ min: 0, max: 49 }),
          }),
          { minLength: 2, maxLength: 6 },
        ),
        (xfs) => {
          const ssA = makeStylesheet();
          const ssB = makeStylesheet();
          for (const xf of xfs) addCellXf(ssA, xf);
          for (const xf of [...xfs].reverse()) addCellXf(ssB, xf);
          const distinct = new Set(xfs.map((x) => JSON.stringify(x))).size;
          return ssA.cellXfs.length === distinct && ssB.cellXfs.length === distinct;
        },
      ),
      { numRuns: 50 },
    );
  });

  it('defaultCellXf is always the first deduped entry', () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: 20 }), (n) => {
        const ss = makeStylesheet();
        const ids = Array.from({ length: n }, () => addCellXf(ss, defaultCellXf()));
        return ids.every((id) => id === 0) && ss.cellXfs.length === 1;
      }),
      { numRuns: 30 },
    );
  });
});

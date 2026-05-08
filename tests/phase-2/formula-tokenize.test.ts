import { describe, expect, it } from 'vitest';
import {
  ARG,
  ARRAY,
  CLOSE,
  ERROR,
  FUNC,
  LITERAL,
  LOGICAL,
  makeOperand,
  makeSeparator,
  makeSubexp,
  NUMBER,
  OP_IN,
  OP_POST,
  OP_PRE,
  OPEN,
  OPERAND,
  PAREN,
  RANGE,
  ROW,
  renderTokens,
  SEP,
  SN_RE,
  STRING_DOUBLE_RE,
  STRING_SINGLE_RE,
  TEXT,
  TokenizerError,
  tokenize,
  WSPACE,
  WSPACE_RE,
} from '../../src/xlsx/formula/tokenizer';

// --- regex tests (mirrors openpyxl test_tokenizer.TestTokenizerRegexes) ----

describe('Tokenizer regexes', () => {
  it.each([
    ['1.0E', true],
    ['1.53321E', true],
    ['9.999E', true],
    ['3E', true],
    ['12E', false],
    ['0.1E', false],
    ['0E', false],
    ['', false],
    ['E', false],
  ] as const)('SN_RE matches %s → %s', (s, ok) => {
    expect(SN_RE.test(s)).toBe(ok);
  });

  it.each([
    [' ', ' '],
    [' *', ' '],
    ['     ', '     '],
    ['     a', '     '],
    ['   ', '   '],
    ['   +', '   '],
    ['', null],
    ['*', null],
  ] as const)('WSPACE_RE on %s → %s', (s, expected) => {
    const m = WSPACE_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect(m?.[0]).toBe(expected);
  });

  it.each([
    ['"spamspamspam"', '"spamspamspam"'],
    ['"this is "" a test "" "', '"this is "" a test "" "'],
    ['""', '""'],
    ['"spam and ""cheese"""+"ignore"', '"spam and ""cheese"""'],
    ['\'"spam and ""cheese"""+"ignore"', null],
    ['"oops ""', null],
  ] as const)('STRING_DOUBLE_RE on %s → %s', (s, expected) => {
    const m = STRING_DOUBLE_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect(m?.[0]).toBe(expected);
  });

  it.each([
    ["'spam and ham'", "'spam and ham'"],
    ["'double'' triple''' quadruple ''''", "'double'' triple'''"],
    ["'sextuple '''''' and septuple''''''' and more", "'sextuple '''''' and septuple'''''''"],
    ["''", "''"],
    ["'oops ''", null],
    ["gunk'hello world'", null],
  ] as const)('STRING_SINGLE_RE on %s → %s', (s, expected) => {
    const m = STRING_SINGLE_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect(m?.[0]).toBe(expected);
  });
});

// --- full-formula parse cases (mirrors openpyxl test_parse) ----------------

type Tup = [string, string, string];

const PARSE_CASES: ReadonlyArray<readonly [string, ReadonlyArray<Tup>]> = [
  [
    '=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))',
    [
      ['IF(', FUNC, OPEN],
      ['A$3', OPERAND, RANGE],
      ['<', OP_IN, ''],
      ['40', OPERAND, NUMBER],
      ['%', OP_POST, ''],
      [',', SEP, ARG],
      ['""', OPERAND, TEXT],
      [',', SEP, ARG],
      ['INDEX(', FUNC, OPEN],
      ['Pipeline!B$4:B$138', OPERAND, RANGE],
      [',', SEP, ARG],
      ['#REF!', OPERAND, ERROR],
      [')', FUNC, CLOSE],
      [')', FUNC, CLOSE],
    ],
  ],
  ["='Summary slices'!$C$3", [["'Summary slices'!$C$3", OPERAND, RANGE]]],
  [
    '=-MAX(Pipeline!AA4:AA138)',
    [
      ['-', OP_PRE, ''],
      ['MAX(', FUNC, OPEN],
      ['Pipeline!AA4:AA138', OPERAND, RANGE],
      [')', FUNC, CLOSE],
    ],
  ],
  [
    '=TEXT(-S7/1000,"$#,##0""M""")',
    [
      ['TEXT(', FUNC, OPEN],
      ['-', OP_PRE, ''],
      ['S7', OPERAND, RANGE],
      ['/', OP_IN, ''],
      ['1000', OPERAND, NUMBER],
      [',', SEP, ARG],
      ['"$#,##0""M"""', OPERAND, TEXT],
      [')', FUNC, CLOSE],
    ],
  ],
  [
    '=IF(A$3<1.3E-8,"",IF(ISNA(\'External Ref\'!K7),"N/A",TEXT(K7*1E+12,"0")&"bp"',
    [
      ['IF(', FUNC, OPEN],
      ['A$3', OPERAND, RANGE],
      ['<', OP_IN, ''],
      ['1.3E-8', OPERAND, NUMBER],
      [',', SEP, ARG],
      ['""', OPERAND, TEXT],
      [',', SEP, ARG],
      ['IF(', FUNC, OPEN],
      ['ISNA(', FUNC, OPEN],
      ["'External Ref'!K7", OPERAND, RANGE],
      [')', FUNC, CLOSE],
      [',', SEP, ARG],
      ['"N/A"', OPERAND, TEXT],
      [',', SEP, ARG],
      ['TEXT(', FUNC, OPEN],
      ['K7', OPERAND, RANGE],
      ['*', OP_IN, ''],
      ['1E+12', OPERAND, NUMBER],
      [',', SEP, ARG],
      ['"0"', OPERAND, TEXT],
      [')', FUNC, CLOSE],
      ['&', OP_IN, ''],
      ['"bp"', OPERAND, TEXT],
    ],
  ],
  [
    '=+IF(A$3<>$B7,"",(MIN(IF({TRUE, FALSE;1,2},A6:B6,$S7))>=LOWER_BOUND)*($BR6>$S72123))',
    [
      ['+', OP_PRE, ''],
      ['IF(', FUNC, OPEN],
      ['A$3', OPERAND, RANGE],
      ['<>', OP_IN, ''],
      ['$B7', OPERAND, RANGE],
      [',', SEP, ARG],
      ['""', OPERAND, TEXT],
      [',', SEP, ARG],
      ['(', PAREN, OPEN],
      ['MIN(', FUNC, OPEN],
      ['IF(', FUNC, OPEN],
      ['{', ARRAY, OPEN],
      ['TRUE', OPERAND, LOGICAL],
      [',', SEP, ARG],
      [' ', WSPACE, ''],
      ['FALSE', OPERAND, LOGICAL],
      [';', SEP, ROW],
      ['1', OPERAND, NUMBER],
      [',', SEP, ARG],
      ['2', OPERAND, NUMBER],
      ['}', ARRAY, CLOSE],
      [',', SEP, ARG],
      ['A6:B6', OPERAND, RANGE],
      [',', SEP, ARG],
      ['$S7', OPERAND, RANGE],
      [')', FUNC, CLOSE],
      [')', FUNC, CLOSE],
      ['>=', OP_IN, ''],
      ['LOWER_BOUND', OPERAND, RANGE],
      [')', PAREN, CLOSE],
      ['*', OP_IN, ''],
      ['(', PAREN, OPEN],
      ['$BR6', OPERAND, RANGE],
      ['>', OP_IN, ''],
      ['$S72123', OPERAND, RANGE],
      [')', PAREN, CLOSE],
      [')', FUNC, CLOSE],
    ],
  ],
  [
    '=(AW$4=$D7)+0%',
    [
      ['(', PAREN, OPEN],
      ['AW$4', OPERAND, RANGE],
      ['=', OP_IN, ''],
      ['$D7', OPERAND, RANGE],
      [')', PAREN, CLOSE],
      ['+', OP_IN, ''],
      ['0', OPERAND, NUMBER],
      ['%', OP_POST, ''],
    ],
  ],
  [
    '=$A:$A,$C:$C',
    [
      ['$A:$A', OPERAND, RANGE],
      [',', OP_IN, ''],
      ['$C:$C', OPERAND, RANGE],
    ],
  ],
  [
    '=3 +1-5',
    [
      ['3', OPERAND, NUMBER],
      [' ', WSPACE, ''],
      ['+', OP_IN, ''],
      ['1', OPERAND, NUMBER],
      ['-', OP_IN, ''],
      ['5', OPERAND, NUMBER],
    ],
  ],
  ['Just text', [['Just text', LITERAL, '']]],
  ['123.456', [['123.456', LITERAL, '']]],
  ['31/12/1999', [['31/12/1999', LITERAL, '']]],
  ['', []],
  [
    '=A1+\nA2',
    [
      ['A1', OPERAND, RANGE],
      ['+', OP_IN, ''],
      ['\n', WSPACE, ''],
      ['A2', OPERAND, RANGE],
    ],
  ],
  ['=R[41]C[2]', [['R[41]C[2]', OPERAND, RANGE]]],
];

describe('tokenize — full formulas', () => {
  for (const [formula, expected] of PARSE_CASES) {
    it(`tokenize(${JSON.stringify(formula)})`, () => {
      const items = tokenize(formula);
      const got: Tup[] = items.map((t) => [t.value, t.type, t.subtype]);
      expect(got).toEqual(expected);
      expect(renderTokens(items)).toBe(formula);
    });
  }
});

// --- error-code parse cases ------------------------------------------------

describe('error-code operands', () => {
  it.each([
    '#NULL!',
    '#DIV/0!',
    '#VALUE!',
    '#REF!',
    '#NAME?',
    '#NUM!',
    '#N/A',
    '#GETTING_DATA',
  ])('=%s tokenizes as a single ERROR operand', (err) => {
    const items = tokenize(`=${err}`);
    expect(items).toHaveLength(1);
    expect(items[0]).toEqual({ value: err, type: OPERAND, subtype: ERROR });
  });

  it('rejects unknown error codes', () => {
    expect(() => tokenize('=#NotAnError')).toThrowError(TokenizerError);
  });

  it('defined-name reference + #REF! is a single RANGE token', () => {
    const items = tokenize('=SUM(MyTable!#REF!)');
    expect(items.map((t) => [t.value, t.type, t.subtype])).toEqual([
      ['SUM(', FUNC, OPEN],
      ['MyTable!#REF!', OPERAND, RANGE],
      [')', FUNC, CLOSE],
    ]);
  });
});

// --- structural error cases -----------------------------------------------

describe('structural errors', () => {
  it.each(['=[unfinished business', '=[un[finished business]'])('rejects unmatched brackets in %s', (formula) => {
    expect(() => tokenize(formula)).toThrowError(TokenizerError);
  });

  it('rejects mismatched ( and { pair', () => {
    expect(() => tokenize('=FUNC(a}')).toThrowError(TokenizerError);
    expect(() => tokenize('=(a}')).toThrowError(TokenizerError);
    expect(() => tokenize('={a)')).toThrowError(TokenizerError);
  });

  it('rejects unterminated string', () => {
    expect(() => tokenize('="unterminated')).toThrowError(TokenizerError);
  });
});

// --- Token factory cases (mirrors openpyxl TestToken) ---------------------

describe('makeOperand subtype detection', () => {
  it.each([
    ['"text"', TEXT],
    ['#REF!', ERROR],
    ['123', NUMBER],
    ['0', NUMBER],
    ['0.123', NUMBER],
    ['.123', NUMBER],
    ['1.234E5', NUMBER],
    ['1E+5', NUMBER],
    ['1.13E-55', NUMBER],
    ['TRUE', LOGICAL],
    ['FALSE', LOGICAL],
    ['A1', RANGE],
    ['ABCD12345', RANGE],
    ["'Hello world'!R123C[-12]", RANGE],
    ["[outside-workbook.xlsx]'A sheet name'!$AB$122", RANGE],
  ])('%s → %s', (value, expected) => {
    const t = makeOperand(value);
    expect(t.type).toBe(OPERAND);
    expect(t.subtype).toBe(expected);
    expect(t.value).toBe(value);
  });
});

describe('makeSubexp / makeSeparator', () => {
  it.each([
    ['{', ARRAY, OPEN],
    ['}', ARRAY, CLOSE],
    ['(', PAREN, OPEN],
    [')', PAREN, CLOSE],
    ['FUNC(', FUNC, OPEN],
  ])('makeSubexp %s → %s/%s', (value, type, sub) => {
    const t = makeSubexp(value);
    expect(t).toEqual({ value, type, subtype: sub });
  });

  it('makeSubexp with func=true forces FUNC type', () => {
    expect(makeSubexp(')', true)).toEqual({ value: ')', type: FUNC, subtype: CLOSE });
    expect(makeSubexp('TEST(', true)).toEqual({ value: 'TEST(', type: FUNC, subtype: OPEN });
  });

  it('makeSeparator distinguishes , vs ;', () => {
    expect(makeSeparator(',')).toEqual({ value: ',', type: SEP, subtype: ARG });
    expect(makeSeparator(';')).toEqual({ value: ';', type: SEP, subtype: ROW });
  });
});

// --- quoted sheet-name in range refs --------------------------------------

describe('quoted sheet name in range ref', () => {
  it.each([
    ["SUM(Inputs!$W$111:'Input 1'!W111)", [["SUM(Inputs!$W$111:'Input 1'!W111)", LITERAL, '']]],
    [
      "=SUM('Inputs 1'!$W$111:'Input 1'!W111)",
      [
        ['SUM(', FUNC, OPEN],
        ["'Inputs 1'!$W$111:'Input 1'!W111", OPERAND, RANGE],
        [')', FUNC, CLOSE],
      ],
    ],
    [
      "=SUM(Inputs!$W$111:'Input 1'!W111)",
      [
        ['SUM(', FUNC, OPEN],
        ["Inputs!$W$111:'Input 1'!W111", OPERAND, RANGE],
        [')', FUNC, CLOSE],
      ],
    ],
    [
      "=SUM(Inputs!$W$111:'Input ''\"1'!W111)",
      [
        ['SUM(', FUNC, OPEN],
        ["Inputs!$W$111:'Input ''\"1'!W111", OPERAND, RANGE],
        [')', FUNC, CLOSE],
      ],
    ],
    [
      '=SUM(Inputs!$W$111:Input1!W111)',
      [
        ['SUM(', FUNC, OPEN],
        ['Inputs!$W$111:Input1!W111', OPERAND, RANGE],
        [')', FUNC, CLOSE],
      ],
    ],
  ] as ReadonlyArray<readonly [string, ReadonlyArray<Tup>]>)('%s', (formula, expected) => {
    const items = tokenize(formula);
    expect(items.map((t) => [t.value, t.type, t.subtype])).toEqual(expected);
    expect(renderTokens(items)).toBe(formula);
  });
});

// --- table structured ref brackets are kept inside RANGE operand ----------

describe('structured table refs', () => {
  it('TableX[[#Data],[COL1]] tokenizes as one RANGE', () => {
    const items = tokenize('=TableX[[#Data],[COL1]]');
    expect(items).toEqual([{ value: 'TableX[[#Data],[COL1]]', type: OPERAND, subtype: RANGE }]);
  });

  it('TableX[[#Data],[COL1]:[COL2]] tokenizes as one RANGE', () => {
    const items = tokenize('=TableX[[#Data],[COL1]:[COL2]]');
    expect(items).toEqual([{ value: 'TableX[[#Data],[COL1]:[COL2]]', type: OPERAND, subtype: RANGE }]);
  });
});

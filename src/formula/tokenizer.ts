// Excel formula tokenizer. TS port of openpyxl/openpyxl/formula/tokenizer.py
// (which itself is a port of Eric Bachtal's JavaScript reference impl). The
// tokenizer is **not** an evaluator — it splits a formula string into a flat
// list of typed tokens that the writer / shared-formula translator consume.
//
// Per the project no-classes rule, the state is kept in a `TokenizerState`
// struct and the parsing rules are free functions that mutate it. The public
// surface is `tokenize(formula)` which returns the resulting token array, plus
// `renderTokens(items)` for the `Tokenizer.render()` equivalent (used by tests
// + array-formula round-trip).

import { OpenXmlError } from '../utils/exceptions';

export class TokenizerError extends OpenXmlError {
  override readonly name = 'TokenizerError';
}

// ---- Token type / subtype constants (mirror openpyxl Token attrs) -----------

export const LITERAL = 'LITERAL';
export const OPERAND = 'OPERAND';
export const FUNC = 'FUNC';
export const ARRAY = 'ARRAY';
export const PAREN = 'PAREN';
export const SEP = 'SEP';
export const OP_PRE = 'OPERATOR-PREFIX';
export const OP_IN = 'OPERATOR-INFIX';
export const OP_POST = 'OPERATOR-POSTFIX';
export const WSPACE = 'WHITE-SPACE';

export type TokenType =
  | typeof LITERAL
  | typeof OPERAND
  | typeof FUNC
  | typeof ARRAY
  | typeof PAREN
  | typeof SEP
  | typeof OP_PRE
  | typeof OP_IN
  | typeof OP_POST
  | typeof WSPACE;

export const TEXT = 'TEXT';
export const NUMBER = 'NUMBER';
export const LOGICAL = 'LOGICAL';
export const ERROR = 'ERROR';
export const RANGE = 'RANGE';
export const OPEN = 'OPEN';
export const CLOSE = 'CLOSE';
export const ARG = 'ARG';
export const ROW = 'ROW';

export type TokenSubtype =
  | ''
  | typeof TEXT
  | typeof NUMBER
  | typeof LOGICAL
  | typeof ERROR
  | typeof RANGE
  | typeof OPEN
  | typeof CLOSE
  | typeof ARG
  | typeof ROW;

export interface Token {
  value: string;
  type: TokenType;
  subtype: TokenSubtype;
}

// ---- Token factories --------------------------------------------------------

const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?$/;

/** Build an OPERAND token, deriving the subtype from the value's shape. */
export function makeOperand(value: string): Token {
  let subtype: TokenSubtype;
  if (value.startsWith('"')) {
    subtype = TEXT;
  } else if (value.startsWith('#')) {
    subtype = ERROR;
  } else if (value === 'TRUE' || value === 'FALSE') {
    subtype = LOGICAL;
  } else if (NUMBER_RE.test(value)) {
    subtype = NUMBER;
  } else {
    subtype = RANGE;
  }
  return { value, type: OPERAND, subtype };
}

/**
 * Build a "subexpression" token (FUNC / PAREN / ARRAY, OPEN or CLOSE).
 *
 * `value` must end with one of `{ } ( )`. If `func` is true, the type is forced
 * to FUNC regardless of the bare-bracket heuristic — mirrors the `func=True`
 * overload in openpyxl `Token.make_subexp`.
 */
export function makeSubexp(value: string, func = false): Token {
  const last = value[value.length - 1];
  if (last !== '{' && last !== '}' && last !== '(' && last !== ')') {
    throw new TokenizerError(`makeSubexp: value must end with { } ( ); got "${value}"`);
  }
  let type: TokenType;
  if (func) {
    type = FUNC;
  } else if (value === '{' || value === '}') {
    type = ARRAY;
  } else if (value === '(' || value === ')') {
    type = PAREN;
  } else {
    type = FUNC;
  }
  const subtype: TokenSubtype = value === ')' || value === '}' ? CLOSE : OPEN;
  return { value, type, subtype };
}

/** Returns a CLOSE token that matches the given OPEN subexpression. */
export function getCloser(open: Token): Token {
  if (open.type !== FUNC && open.type !== ARRAY && open.type !== PAREN) {
    throw new TokenizerError(`getCloser: expected FUNC/ARRAY/PAREN OPEN; got ${open.type}`);
  }
  if (open.subtype !== OPEN) {
    throw new TokenizerError(`getCloser: expected OPEN subtype; got ${open.subtype}`);
  }
  const value = open.type === ARRAY ? '}' : ')';
  return makeSubexp(value, open.type === FUNC);
}

/** Build a SEP token (',' → ARG, ';' → ROW). */
export function makeSeparator(value: ',' | ';'): Token {
  return { value, type: SEP, subtype: value === ',' ? ARG : ROW };
}

// ---- Public regex constants (mirrors the openpyxl class attrs) --------------

/** Scientific-notation guard: matches a coefficient ending in `[Ee]`. */
export const SN_RE = /^[1-9](?:\.[0-9]+)?[Ee]$/;

/** Whitespace run (space + LF, per Excel formula whitespace rules). */
export const WSPACE_RE = /^[ \n]+/;

/** "..."-delimited string literal. Internal `""` is an escaped quote. */
export const STRING_DOUBLE_RE = /^"(?:[^"]*"")*[^"]*"(?!")/;
/** '...'-delimited link / sheet name. Internal `''` is an escaped apostrophe. */
export const STRING_SINGLE_RE = /^'(?:[^']*'')*[^']*'(?!')/;

const ERROR_CODES: readonly string[] = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
];

/** Each of these characters terminates the operand token currently being built. */
const TOKEN_ENDERS = new Set(',;}) +-*/^&=><%'.split(''));

// ---- Tokenizer state --------------------------------------------------------

interface TokenizerState {
  formula: string;
  items: Token[];
  /** Stack of OPEN tokens for ARRAY / FUNC / PAREN balance tracking. */
  tokenStack: Token[];
  /** How many chars consumed so far. */
  offset: number;
  /** In-progress operand text, char by char. */
  token: string[];
}

function createState(formula: string): TokenizerState {
  return { formula, items: [], tokenStack: [], offset: 0, token: [] };
}

function saveToken(s: TokenizerState): void {
  if (s.token.length > 0) {
    s.items.push(makeOperand(s.token.join('')));
    s.token.length = 0;
  }
}

function assertEmptyToken(s: TokenizerState, canFollow = ''): void {
  if (s.token.length === 0) return;
  const last = s.token[s.token.length - 1] as string;
  if (canFollow.indexOf(last) === -1) {
    throw new TokenizerError(`Unexpected character at position ${s.offset} in '${s.formula}'`);
  }
}

// ---- Per-rule consumers (return # chars consumed, do not mutate offset) -----

function parseStringLiteral(s: TokenizerState): number {
  assertEmptyToken(s, ':');
  const delim = s.formula[s.offset];
  const regex = delim === '"' ? STRING_DOUBLE_RE : STRING_SINGLE_RE;
  const tail = s.formula.slice(s.offset);
  const m = regex.exec(tail);
  if (m === null) {
    const subtype = delim === '"' ? 'string' : 'link';
    throw new TokenizerError(`Reached end of formula while parsing ${subtype} in ${s.formula}`);
  }
  const match = m[0];
  if (delim === '"') {
    s.items.push(makeOperand(match));
  } else {
    s.token.push(match);
  }
  return match.length;
}

function parseBrackets(s: TokenizerState): number {
  // Walks the bracketed run starting at offset, balancing nested `[`/`]`.
  // Tokens inside brackets are kept verbatim and stitched onto the in-progress
  // operand (e.g. structured table refs `Table1[[#Data],[Col]]`).
  const tail = s.formula.slice(s.offset);
  let openCount = 0;
  for (let i = 0; i < tail.length; i++) {
    const ch = tail[i];
    if (ch === '[') openCount += 1;
    else if (ch === ']') {
      openCount -= 1;
      if (openCount === 0) {
        const outerRight = i + 1;
        s.token.push(tail.slice(0, outerRight));
        return outerRight;
      }
    }
  }
  throw new TokenizerError(`Encountered unmatched '[' in ${s.formula}`);
}

function parseError(s: TokenizerState): number {
  assertEmptyToken(s, '!');
  const sub = s.formula.slice(s.offset);
  for (const err of ERROR_CODES) {
    if (sub.startsWith(err)) {
      s.items.push(makeOperand(s.token.join('') + err));
      s.token.length = 0;
      return err.length;
    }
  }
  throw new TokenizerError(`Invalid error code at position ${s.offset} in '${s.formula}'`);
}

function parseWhitespace(s: TokenizerState): number {
  const ch = s.formula[s.offset] as string;
  s.items.push({ value: ch, type: WSPACE, subtype: '' });
  const tail = s.formula.slice(s.offset);
  const m = WSPACE_RE.exec(tail);
  return m ? m[0].length : 1;
}

function parseOperator(s: TokenizerState): number {
  const two = s.formula.slice(s.offset, s.offset + 2);
  if (two === '>=' || two === '<=' || two === '<>') {
    s.items.push({ value: two, type: OP_IN, subtype: '' });
    return 2;
  }
  const ch = s.formula[s.offset] as string;
  let token: Token;
  if (ch === '%') {
    token = { value: '%', type: OP_POST, subtype: '' };
  } else if ('*/^&=><'.indexOf(ch) !== -1) {
    token = { value: ch, type: OP_IN, subtype: '' };
  } else {
    // ch is in '+-' from here on
    if (s.items.length === 0) {
      token = { value: ch, type: OP_PRE, subtype: '' };
    } else {
      // Find the most recent non-whitespace token to decide infix vs prefix.
      let prev: Token | null = null;
      for (let i = s.items.length - 1; i >= 0; i--) {
        const t = s.items[i] as Token;
        if (t.type !== WSPACE) {
          prev = t;
          break;
        }
      }
      const isInfix = prev !== null && (prev.subtype === CLOSE || prev.type === OP_POST || prev.type === OPERAND);
      token = { value: ch, type: isInfix ? OP_IN : OP_PRE, subtype: '' };
    }
  }
  s.items.push(token);
  return 1;
}

function parseOpener(s: TokenizerState): number {
  const ch = s.formula[s.offset] as string;
  let token: Token;
  if (ch === '{') {
    assertEmptyToken(s);
    token = makeSubexp('{');
  } else if (s.token.length > 0) {
    const value = `${s.token.join('')}(`;
    s.token.length = 0;
    token = makeSubexp(value);
  } else {
    token = makeSubexp('(');
  }
  s.items.push(token);
  s.tokenStack.push(token);
  return 1;
}

function parseCloser(s: TokenizerState): number {
  const opener = s.tokenStack.pop();
  if (!opener) {
    throw new TokenizerError(`Unmatched closer in '${s.formula}'`);
  }
  const closer = getCloser(opener);
  if (closer.value !== s.formula[s.offset]) {
    throw new TokenizerError(`Mismatched ( and { pair in '${s.formula}'`);
  }
  s.items.push(closer);
  return 1;
}

function parseSeparator(s: TokenizerState): number {
  const ch = s.formula[s.offset] as string;
  let token: Token;
  if (ch === ';') {
    token = makeSeparator(';');
  } else {
    // ','. Range-union if the enclosing context is a PAREN or there is none
    // (top-level), otherwise an arg separator.
    const top = s.tokenStack[s.tokenStack.length - 1];
    if (!top || top.type === PAREN) {
      token = { value: ',', type: OP_IN, subtype: '' };
    } else {
      token = makeSeparator(',');
    }
  }
  s.items.push(token);
  return 1;
}

function checkScientificNotation(s: TokenizerState): boolean {
  const ch = s.formula[s.offset];
  if ((ch === '+' || ch === '-') && s.token.length >= 1 && SN_RE.test(s.token.join(''))) {
    s.token.push(ch);
    s.offset += 1;
    return true;
  }
  return false;
}

// ---- Public API -------------------------------------------------------------

/**
 * Tokenize a formula string into a flat list of `Token`s.
 *
 * Mirrors `openpyxl.formula.tokenizer.Tokenizer(formula).items`. A formula that
 * does not start with `=` becomes a single LITERAL token. Empty input yields
 * the empty array.
 */
export function tokenize(formula: string): Token[] {
  const s = createState(formula);
  if (formula.length === 0) return s.items;
  if (formula[0] !== '=') {
    s.items.push({ value: formula, type: LITERAL, subtype: '' });
    return s.items;
  }
  s.offset = 1;
  const len = formula.length;
  while (s.offset < len) {
    if (checkScientificNotation(s)) continue;
    const ch = formula[s.offset] as string;
    if (TOKEN_ENDERS.has(ch)) saveToken(s);
    let consumed: number | null = null;
    switch (ch) {
      case '"':
      case "'":
        consumed = parseStringLiteral(s);
        break;
      case '[':
        consumed = parseBrackets(s);
        break;
      case '#':
        consumed = parseError(s);
        break;
      case ' ':
      case '\n':
        consumed = parseWhitespace(s);
        break;
      case '+':
      case '-':
      case '*':
      case '/':
      case '^':
      case '&':
      case '=':
      case '>':
      case '<':
      case '%':
        consumed = parseOperator(s);
        break;
      case '{':
      case '(':
        consumed = parseOpener(s);
        break;
      case ')':
      case '}':
        consumed = parseCloser(s);
        break;
      case ';':
      case ',':
        consumed = parseSeparator(s);
        break;
      default:
        s.token.push(ch);
        consumed = 1;
    }
    s.offset += consumed;
  }
  saveToken(s);
  return s.items;
}

/** Render a token list back into the original formula string. */
export function renderTokens(items: ReadonlyArray<Token>): string {
  if (items.length === 0) return '';
  const first = items[0] as Token;
  if (first.type === LITERAL) return first.value;
  let out = '=';
  for (const t of items) out += t.value;
  return out;
}

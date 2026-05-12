import { describe, expect, it } from 'vitest';
import { utf8ByteLength } from '../../src/utils/utf8';

const encoder = new TextEncoder();
const oracle = (s: string): number => encoder.encode(s).byteLength;

describe('utf8ByteLength', () => {
  it('matches TextEncoder for ASCII', () => {
    const s = 'hello world';
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('matches TextEncoder for 2-byte codepoints', () => {
    const s = 'café — naïve';
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('matches TextEncoder for 3-byte BMP codepoints (CJK)', () => {
    const s = '日本語のテキスト ABC 中文';
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('matches TextEncoder for valid surrogate pairs (4-byte codepoints)', () => {
    // U+1F600 (😀) — encoded as the surrogate pair D83D DE00 in UTF-16.
    const s = '😀abc🚀';
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('handles a lone high surrogate at end of string like TextEncoder', () => {
    const s = `abc${String.fromCharCode(0xd83d)}`;
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('handles a high surrogate followed by a non-low-surrogate without swallowing the next code unit', () => {
    // High surrogate immediately followed by a BMP character. TextEncoder
    // emits U+FFFD (3 bytes) for the unpaired surrogate, then the BMP
    // character on its own. The previous implementation would swallow the
    // 'X' entirely.
    const s = `${String.fromCharCode(0xd83d)}X`;
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('handles a lone low surrogate', () => {
    const s = `${String.fromCharCode(0xdc00)}Y`;
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });

  it('handles back-to-back high surrogates (each unpaired)', () => {
    const s = `${String.fromCharCode(0xd83d)}${String.fromCharCode(0xd83d)}`;
    expect(utf8ByteLength(s)).toBe(oracle(s));
  });
});

// Tests for the public `insertRichText(rt, index, insertion)` helper.

import { describe, expect, it } from 'vitest';
import {
  insertRichText,
  makeRichText,
  richText,
  richTextToString,
} from '../../src/cell';

describe('insertRichText', () => {
  it('inserts a string in the middle of the run text', () => {
    const rt = makeRichText([{ text: 'helloworld' }]);
    const out = insertRichText(rt, 5, ' ');
    expect(richTextToString(out)).toBe('hello world');
  });

  it('inserts at index 0 (prepend)', () => {
    const rt = makeRichText([{ text: 'world' }]);
    const out = insertRichText(rt, 0, 'hello ');
    expect(richTextToString(out)).toBe('hello world');
  });

  it('inserts at the end (index = length, append)', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    const out = insertRichText(rt, 5, '!');
    expect(richTextToString(out)).toBe('hello!');
  });

  it('preserves font on a RichText insertion and accepts negative index', () => {
    const rt = makeRichText([{ text: 'hello!' }]);
    const out = insertRichText(rt, -1, richText(' world', { i: true }));
    expect(richTextToString(out)).toBe('hello world!');
    const italicRun = out.find((r) => r.text === ' world');
    expect(italicRun?.font).toEqual({ i: true });
  });
});

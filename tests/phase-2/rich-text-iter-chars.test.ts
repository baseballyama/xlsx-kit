// Tests for the public `iterRichTextChars(rt)` generator.

import { describe, expect, it } from 'vitest';
import { iterRichTextChars, makeRichText } from '../../src/cell';

describe('iterRichTextChars', () => {
  it('yields each character of a single styled run', () => {
    const rt = makeRichText([{ text: 'abc', font: { b: true } }]);
    const items = [...iterRichTextChars(rt)];
    expect(items).toEqual([
      { char: 'a', font: { b: true }, index: 0 },
      { char: 'b', font: { b: true }, index: 1 },
      { char: 'c', font: { b: true }, index: 2 },
    ]);
  });

  it('switches font as the iterator crosses run boundaries', () => {
    const rt = makeRichText([
      { text: 'A', font: { b: true } },
      { text: 'B' },
      { text: 'C', font: { i: true } },
    ]);
    const items = [...iterRichTextChars(rt)];
    expect(items).toEqual([
      { char: 'A', font: { b: true }, index: 0 },
      { char: 'B', font: undefined, index: 1 },
      { char: 'C', font: { i: true }, index: 2 },
    ]);
  });

  it('yields nothing for an empty RichText', () => {
    expect([...iterRichTextChars(makeRichText([]))]).toEqual([]);
  });
});

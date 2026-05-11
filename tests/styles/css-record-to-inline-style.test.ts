// Tests for cssRecordToInlineStyle — Record<string,string> → "k: v; …".

import { describe, expect, it } from 'vitest';
import { cssRecordToInlineStyle } from '../../src/utils/css';

describe('cssRecordToInlineStyle', () => {
  it('returns empty string for {} / undefined', () => {
    expect(cssRecordToInlineStyle({})).toBe('');
    expect(cssRecordToInlineStyle(undefined)).toBe('');
  });

  it('serialises a single property without trailing semicolon', () => {
    expect(cssRecordToInlineStyle({ color: '#FF0000' })).toBe('color: #FF0000');
  });

  it('alphabetises multi-property output', () => {
    expect(
      cssRecordToInlineStyle({
        'font-weight': 'bold',
        color: '#000000',
        'background-color': '#FFFF00',
      }),
    ).toBe('background-color: #FFFF00; color: #000000; font-weight: bold');
  });

  it('skips empty-string values (treated as unset)', () => {
    expect(cssRecordToInlineStyle({ color: '', 'font-weight': 'bold' })).toBe('font-weight: bold');
  });

  it('drops values containing `;` (defensive against attribute-injection)', () => {
    expect(
      cssRecordToInlineStyle({ color: 'red; -webkit-something: bad', 'font-weight': 'bold' }),
    ).toBe('font-weight: bold');
  });
});

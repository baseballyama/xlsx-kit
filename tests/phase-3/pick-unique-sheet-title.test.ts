// Tests for pickUniqueSheetTitle.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  pickUniqueSheetTitle,
} from '../../src/workbook/workbook';

describe('pickUniqueSheetTitle', () => {
  it('returns the base verbatim when free', () => {
    const wb = createWorkbook();
    expect(pickUniqueSheetTitle(wb, 'Sheet1')).toBe('Sheet1');
  });

  it('appends " (2)" / " (3)" until free', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    expect(pickUniqueSheetTitle(wb, 'Sheet1')).toBe('Sheet1 (2)');
    addWorksheet(wb, 'Sheet1 (2)');
    expect(pickUniqueSheetTitle(wb, 'Sheet1')).toBe('Sheet1 (3)');
  });

  it('truncates the base when base+suffix would exceed 31 chars', () => {
    const wb = createWorkbook();
    const base = 'x'.repeat(31);
    addWorksheet(wb, base);
    const out = pickUniqueSheetTitle(wb, base);
    expect(out.length).toBeLessThanOrEqual(31);
    expect(out).toBe(`${'x'.repeat(27)} (2)`);
  });

  it('throws when base itself is invalid', () => {
    const wb = createWorkbook();
    expect(() => pickUniqueSheetTitle(wb, '')).toThrow(/not a valid sheet title/);
    expect(() => pickUniqueSheetTitle(wb, 'Sheet:1')).toThrow(/not a valid sheet title/);
  });

  it('returns a value that passes validateSheetTitle (uniqueified result fits Excel rules)', async () => {
    const { validateSheetTitle } = await import('../../src/workbook/workbook');
    const wb = createWorkbook();
    const base = 'x'.repeat(31);
    addWorksheet(wb, base);
    const out = pickUniqueSheetTitle(wb, base);
    expect(validateSheetTitle(out)).toBeUndefined();
  });
});

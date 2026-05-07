// Tests for hasWorksheet — worksheet-only contains predicate.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  hasWorksheet,
} from '../../src/workbook/workbook';

describe('hasWorksheet', () => {
  it('returns true when a worksheet with the title exists', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(hasWorksheet(wb, 'Data')).toBe(true);
  });

  it('returns false when only a chartsheet (not a worksheet) shares the title', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Shared');
    expect(hasWorksheet(wb, 'Shared')).toBe(false);
  });

  it('returns false when the title is not present at all', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Other');
    expect(hasWorksheet(wb, 'Missing')).toBe(false);
  });

  it('returns false for an empty workbook', () => {
    expect(hasWorksheet(createWorkbook(), 'anything')).toBe(false);
  });
});

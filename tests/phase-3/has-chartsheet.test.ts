// Tests for hasChartsheet — chartsheet-only contains predicate.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  hasChartsheet,
} from '../../src/xlsx/workbook/workbook';

describe('hasChartsheet', () => {
  it('returns true when a chartsheet with the title exists', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    expect(hasChartsheet(wb, 'Chart1')).toBe(true);
  });

  it('returns false when the title is not present', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    expect(hasChartsheet(wb, 'Missing')).toBe(false);
  });

  it('returns false when only a worksheet (not a chartsheet) shares the title', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Shared');
    expect(hasChartsheet(wb, 'Shared')).toBe(false);
  });

  it('returns false for an empty workbook', () => {
    expect(hasChartsheet(createWorkbook(), 'anything')).toBe(false);
  });
});

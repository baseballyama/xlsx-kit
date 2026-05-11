// Tests for hasSheet — workbook contains predicate.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  hasSheet,
} from '../../src/workbook/workbook';

describe('hasSheet', () => {
  it('returns true when a worksheet with the title exists', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(hasSheet(wb, 'Data')).toBe(true);
  });

  it('returns false when the title is not present', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Other');
    expect(hasSheet(wb, 'Missing')).toBe(false);
  });

  it('finds chartsheets by title too', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    expect(hasSheet(wb, 'Chart1')).toBe(true);
  });

  it('returns false for an empty workbook', () => {
    expect(hasSheet(createWorkbook(), 'anything')).toBe(false);
  });
});

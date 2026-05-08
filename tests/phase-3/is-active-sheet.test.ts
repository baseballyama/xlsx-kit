// Tests for isActiveSheet — title-vs-active-sheet predicate.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  isActiveSheet,
  setActiveSheet,
} from '../../src/xlsx/workbook/workbook';

describe('isActiveSheet', () => {
  it('returns true when the title matches the active sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    expect(isActiveSheet(wb, 'first')).toBe(true);
  });

  it('returns false when the title does not match the active sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    expect(isActiveSheet(wb, 'second')).toBe(false);
  });

  it('reflects setActiveSheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    setActiveSheet(wb, 'second');
    expect(isActiveSheet(wb, 'first')).toBe(false);
    expect(isActiveSheet(wb, 'second')).toBe(true);
  });

  it('returns false for an empty workbook', () => {
    expect(isActiveSheet(createWorkbook(), 'anything')).toBe(false);
  });
});

// Tests for setActiveCell / setSelectedRange.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  setActiveCell,
  setSelectedRange,
  type Worksheet,
} from '../../src/worksheet/worksheet';

describe('setActiveCell', () => {
  it('lazily creates a Selection on the primary view', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'B5');
    expect(ws.views[0]?.selection?.activeCell).toBe('B5');
    expect(ws.views[0]?.selection?.sqref).toBe('B5');
  });

  it('updates activeCell + sqref together when sqref tracked the previous activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'A1');
    setActiveCell(ws, 'C3');
    expect(ws.views[0]?.selection?.activeCell).toBe('C3');
    expect(ws.views[0]?.selection?.sqref).toBe('C3');
  });

  it('preserves an explicitly-set sqref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'A1:D10');
    setActiveCell(ws, 'B2');
    expect(ws.views[0]?.selection?.activeCell).toBe('B2');
    expect(ws.views[0]?.selection?.sqref).toBe('A1:D10');
  });
});

describe('setSelectedRange', () => {
  it('single range sets sqref + derives activeCell from top-left when missing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'B2:D5');
    expect(ws.views[0]?.selection?.sqref).toBe('B2:D5');
    expect(ws.views[0]?.selection?.activeCell).toBe('B2');
  });

  it('multi-range sqref keeps the first ref as activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'A1 C3:D4');
    expect(ws.views[0]?.selection?.activeCell).toBe('A1');
    expect(ws.views[0]?.selection?.sqref).toBe('A1 C3:D4');
  });

  it('does not overwrite an existing activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'E10');
    setSelectedRange(ws, 'A1:B2');
    expect(ws.views[0]?.selection?.activeCell).toBe('E10');
    expect(ws.views[0]?.selection?.sqref).toBe('A1:B2');
  });
});

describe('selection round-trip', () => {
  it('selection survives save → load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sel');
    setSelectedRange(ws, 'A1:C5');
    setActiveCell(ws, 'B3');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.views[0]?.selection?.activeCell).toBe('B3');
    expect(ws2.views[0]?.selection?.sqref).toBe('A1:C5');
  });
});

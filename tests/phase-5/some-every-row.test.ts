// Tests for someRow / everyRow — header-driven Array.some / .every for rows.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { everyRow, setCell, someRow } from '../../src/worksheet/worksheet';

describe('someRow', () => {
  it('returns true when at least one row matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'target');
    setCell(ws, 4, 1, 'c');
    expect(someRow(ws, 'A1:A4', (r) => r['k'] === 'target')).toBe(true);
  });

  it('returns false when no row matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(someRow(ws, 'A1:A3', () => false)).toBe(false);
  });

  it('returns false for an empty data area (header only)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(someRow(ws, 'A1:A1', () => true)).toBe(false);
  });

  it('short-circuits at the first match', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'match');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    let calls = 0;
    expect(
      someRow(ws, 'A1:A4', () => {
        calls++;
        return true;
      }),
    ).toBe(true);
    expect(calls).toBe(1);
  });
});

describe('everyRow', () => {
  it('returns true when every row matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    expect(everyRow(ws, 'A1:A4', (r) => typeof r['n'] === 'number')).toBe(true);
  });

  it('returns false at the first failing row (short-circuit)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 'oops');
    setCell(ws, 4, 1, 3);
    let calls = 0;
    const result = everyRow(ws, 'A1:A4', (r) => {
      calls++;
      return typeof r['n'] === 'number';
    });
    expect(result).toBe(false);
    expect(calls).toBe(2); // 1 (pass) + 'oops' (fail) → stops
  });

  it('returns true vacuously for an empty data area', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(everyRow(ws, 'A1:A1', () => false)).toBe(true);
  });
});

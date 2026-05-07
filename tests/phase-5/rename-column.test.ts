// Tests for renameColumn — header-row rename in place.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { readRangeAsObjects, renameColumn, setCell } from '../../src/worksheet/worksheet';

describe('renameColumn', () => {
  it('rewrites the header cell from oldName to newName', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 30);
    renameColumn(ws, 'A1:B2', 'age', 'years');
    expect(readRangeAsObjects(ws, 'A1:B2')).toEqual([{ name: 'Alice', years: 30 }]);
  });

  it('does not touch any data cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'old');
    setCell(ws, 2, 1, 'old'); // data row also has the value 'old' — must remain
    renameColumn(ws, 'A1:A2', 'old', 'new');
    expect(ws.rows.get(1)?.get(1)?.value).toBe('new');
    expect(ws.rows.get(2)?.get(1)?.value).toBe('old');
  });

  it('throws when oldName is not in the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    expect(() => renameColumn(ws, 'A1:A1', 'missing', 'newname')).toThrow(/missing/);
  });

  it('throws when newName already exists in the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'role');
    expect(() => renameColumn(ws, 'A1:B1', 'name', 'role')).toThrow(/already exists/);
  });

  it('no-op when oldName equals newName', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    renameColumn(ws, 'A1:A1', 'name', 'name');
    expect(ws.rows.get(1)?.get(1)?.value).toBe('name');
  });
});

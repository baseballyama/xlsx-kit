// Tests for the typed <scenarios> model. Per
// docs/plan/13-full-excel-coverage.md §B7 (Scenario Manager).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { parseMultiCellRange } from '../../src/xlsx/worksheet/cell-range';
import { makeScenario, makeScenarioInputCell, makeScenarioList } from '../../src/xlsx/worksheet/scenarios';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet scenarios round-trip', () => {
  it('preserves a scenario list with two scenarios + multiple inputCells', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'rate');
    setCell(ws, 1, 2, 'value');
    setCell(ws, 2, 1, 0.05);

    ws.scenarios = makeScenarioList({
      current: 0,
      show: 0,
      sqref: parseMultiCellRange('A2:A2'),
      scenarios: [
        makeScenario({
          name: 'BaseCase',
          user: 'qa',
          comment: 'default rate',
          locked: false,
          hidden: false,
          inputCells: [makeScenarioInputCell({ ref: 'A2', val: '0.05', numFmtId: 9 })],
        }),
        makeScenario({
          name: 'Optimistic',
          inputCells: [
            makeScenarioInputCell({ ref: 'A2', val: '0.08', numFmtId: 9 }),
            makeScenarioInputCell({ ref: 'B2', val: '1000', numFmtId: 0 }),
          ],
        }),
      ],
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const sl = ws2.scenarios;
    expect(sl).toBeDefined();
    expect(sl?.current).toBe(0);
    expect(sl?.show).toBe(0);
    expect(sl?.scenarios.length).toBe(2);
    expect(sl?.scenarios[0]?.name).toBe('BaseCase');
    expect(sl?.scenarios[0]?.user).toBe('qa');
    expect(sl?.scenarios[0]?.comment).toBe('default rate');
    expect(sl?.scenarios[0]?.locked).toBe(false);
    expect(sl?.scenarios[0]?.inputCells[0]?.ref).toBe('A2');
    expect(sl?.scenarios[0]?.inputCells[0]?.val).toBe('0.05');
    expect(sl?.scenarios[0]?.inputCells[0]?.numFmtId).toBe(9);
    expect(sl?.scenarios[1]?.name).toBe('Optimistic');
    expect(sl?.scenarios[1]?.inputCells.length).toBe(2);
    expect(sl?.scenarios[1]?.inputCells[1]?.val).toBe('1000');
  });

  it('emits no <scenarios/> when undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.scenarios).toBeUndefined();
  });
});
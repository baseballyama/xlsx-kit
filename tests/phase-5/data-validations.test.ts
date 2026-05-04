import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { parseMultiCellRange } from '../../src/worksheet/cell-range';
import { makeDataValidation } from '../../src/worksheet/data-validations';
import { addDataValidation, removeDataValidations, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('makeDataValidation + addDataValidation', () => {
  it('builds a list-type validator and lets it round-trip the API', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addDataValidation(
      ws,
      makeDataValidation({
        type: 'list',
        sqref: parseMultiCellRange('A1:A10'),
        formula1: '"red,green,blue"',
        allowBlank: true,
        showInputMessage: true,
      }),
    );
    expect(ws.dataValidations.length).toBe(1);
    expect(dv.type).toBe('list');
    expect(dv.formula1).toBe('"red,green,blue"');
    expect(dv.sqref.ranges.length).toBe(1);
  });

  it('removeDataValidations filters by predicate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    addDataValidation(ws, makeDataValidation({ type: 'list', sqref: parseMultiCellRange('A1:A5'), formula1: '"x,y"' }));
    addDataValidation(
      ws,
      makeDataValidation({
        type: 'whole',
        sqref: parseMultiCellRange('B1:B5'),
        formula1: '0',
        operator: 'greaterThan',
      }),
    );
    expect(removeDataValidations(ws, (dv) => dv.type === 'whole')).toBe(1);
    expect(ws.dataValidations.length).toBe(1);
    expect(ws.dataValidations[0]?.type).toBe('list');
  });
});

describe('dataValidations round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves a list-type dropdown', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    addDataValidation(
      ws,
      makeDataValidation({
        type: 'list',
        sqref: parseMultiCellRange('A1:A10'),
        formula1: '"alpha,beta,gamma"',
        allowBlank: true,
        showDropDown: false,
        showInputMessage: true,
        showErrorMessage: true,
        promptTitle: 'Pick one',
        prompt: 'Choose alpha, beta, or gamma',
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.dataValidations.length).toBe(1);
    const dv = ws2.dataValidations[0];
    expect(dv?.type).toBe('list');
    expect(dv?.formula1).toBe('"alpha,beta,gamma"');
    expect(dv?.allowBlank).toBe(true);
    expect(dv?.showInputMessage).toBe(true);
    expect(dv?.promptTitle).toBe('Pick one');
  });

  it('preserves whole-number range with operator + formula1/formula2', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    addDataValidation(
      ws,
      makeDataValidation({
        type: 'whole',
        sqref: parseMultiCellRange('B1:B100'),
        operator: 'between',
        formula1: '1',
        formula2: '100',
        errorStyle: 'warning',
        errorTitle: 'Out of range',
        error: 'Pick a value 1-100',
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const dv = ws2.dataValidations[0];
    expect(dv?.type).toBe('whole');
    expect(dv?.operator).toBe('between');
    expect(dv?.formula1).toBe('1');
    expect(dv?.formula2).toBe('100');
    expect(dv?.errorStyle).toBe('warning');
    expect(dv?.errorTitle).toBe('Out of range');
  });

  it('preserves custom-type with a formula referencing other cells', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    addDataValidation(
      ws,
      makeDataValidation({
        type: 'custom',
        sqref: parseMultiCellRange('C2:C5'),
        formula1: 'AND(LEN(C2)<10,EXACT(LOWER(C2),C2))',
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.dataValidations[0]?.type).toBe('custom');
    expect(ws2.dataValidations[0]?.formula1).toBe('AND(LEN(C2)<10,EXACT(LOWER(C2),C2))');
  });

  it('preserves a multi-range sqref', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    addDataValidation(
      ws,
      makeDataValidation({
        type: 'list',
        sqref: parseMultiCellRange('A1:A5 D1:D5'),
        formula1: '"yes,no"',
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const dv = ws2.dataValidations[0];
    expect(dv?.sqref.ranges.length).toBe(2);
    expect(dv?.sqref.ranges[0]?.maxCol).toBe(1);
    expect(dv?.sqref.ranges[1]?.maxCol).toBe(4);
  });

  it('omits the <dataValidations> block when none are set', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoVal');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('<dataValidations');
  });
});

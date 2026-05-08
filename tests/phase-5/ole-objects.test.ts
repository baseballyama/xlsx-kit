// Tests for the typed worksheet <oleObjects> + <controls> models.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeFormControl, makeOleObject } from '../../src/xlsx/worksheet/ole-objects';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet oleObjects round-trip', () => {
  it('preserves multiple oleObject entries with their typed top-level attrs', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'O');
    setCell(ws, 1, 1, 1);
    ws.oleObjects.push(
      makeOleObject({
        shapeId: 1025,
        rId: 'rId4',
        progId: 'Word.Document.12',
        dvAspect: 'DVASPECT_CONTENT',
        oleUpdate: 'OLEUPDATE_ONCALL',
        autoLoad: false,
      }),
    );
    ws.oleObjects.push(
      makeOleObject({
        shapeId: 1026,
        rId: 'rId5',
        progId: 'Equation.3',
        dvAspect: 'DVASPECT_ICON',
        link: '0Equation.DSMT4\t9',
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.oleObjects.length).toBe(2);
    expect(ws2.oleObjects[0]?.shapeId).toBe(1025);
    expect(ws2.oleObjects[0]?.progId).toBe('Word.Document.12');
    expect(ws2.oleObjects[0]?.dvAspect).toBe('DVASPECT_CONTENT');
    expect(ws2.oleObjects[0]?.oleUpdate).toBe('OLEUPDATE_ONCALL');
    expect(ws2.oleObjects[0]?.autoLoad).toBe(false);
    expect(ws2.oleObjects[1]?.shapeId).toBe(1026);
    expect(ws2.oleObjects[1]?.dvAspect).toBe('DVASPECT_ICON');
    expect(ws2.oleObjects[1]?.link).toBe('0Equation.DSMT4\t9');
  });

  it('emits no <oleObjects/> when empty', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.oleObjects.length).toBe(0);
  });
});

describe('worksheet controls round-trip', () => {
  it('preserves form-control entries', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setCell(ws, 1, 1, 1);
    ws.controls.push(makeFormControl({ shapeId: 2049, rId: 'rId10', name: 'CheckBox1' }));
    ws.controls.push(makeFormControl({ shapeId: 2050, rId: 'rId11', name: 'SpinButton1' }));

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.controls.length).toBe(2);
    expect(ws2.controls[0]?.shapeId).toBe(2049);
    expect(ws2.controls[0]?.name).toBe('CheckBox1');
    expect(ws2.controls[1]?.shapeId).toBe(2050);
  });

  it('emits no <controls/> when empty', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.controls.length).toBe(0);
  });
});
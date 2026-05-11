import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeHyperlink } from '../../src/worksheet/hyperlinks';
import { getHyperlink, removeHyperlink, setCell, setHyperlink, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('setHyperlink / getHyperlink / removeHyperlink', () => {
  it('rejects when neither target nor location is set', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    expect(() => setHyperlink(ws, 'A1', {})).toThrowError(/target.*location/);
  });

  it('makeHyperlink throws OpenXmlSchemaError when ref is empty', () => {
    expect(() => makeHyperlink({ ref: '', target: 'https://example.com' })).toThrow(OpenXmlSchemaError);
  });

  it('stores an external link with target', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setHyperlink(ws, 'A1', { target: 'https://example.com', tooltip: 'click me' });
    expect(getHyperlink(ws, 'A1')?.target).toBe('https://example.com');
    expect(getHyperlink(ws, 'A1')?.tooltip).toBe('click me');
  });

  it('stores an internal jump with location only', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setHyperlink(ws, 'B2', { location: "'Sheet 2'!A1" });
    expect(getHyperlink(ws, 'B2')?.location).toBe("'Sheet 2'!A1");
    expect(getHyperlink(ws, 'B2')?.target).toBeUndefined();
  });

  it('replaces an existing entry on the same ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setHyperlink(ws, 'A1', { target: 'https://a.example' });
    setHyperlink(ws, 'A1', { target: 'https://b.example' });
    expect(ws.hyperlinks.length).toBe(1);
    expect(getHyperlink(ws, 'A1')?.target).toBe('https://b.example');
  });

  it('removeHyperlink drops the entry; second call returns false', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setHyperlink(ws, 'A1', { target: 'https://x.example' });
    expect(removeHyperlink(ws, 'A1')).toBe(true);
    expect(removeHyperlink(ws, 'A1')).toBe(false);
    expect(getHyperlink(ws, 'A1')).toBeUndefined();
  });
});

describe('hyperlinks round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves external URL via worksheet rels', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setCell(ws, 1, 1, 'go to example');
    setHyperlink(ws, 'A1', {
      target: 'https://example.com/path',
      tooltip: 'open the example site',
      display: 'example',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);

    const link = getHyperlink(ws2, 'A1');
    expect(link?.target).toBe('https://example.com/path');
    expect(link?.tooltip).toBe('open the example site');
    expect(link?.display).toBe('example');
    expect(link?.rId).toMatch(/^rId\d+$/);
  });

  it('preserves internal jumps without allocating a rels entry', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    addWorksheet(wb, 'Other');
    setHyperlink(ws, 'C3', { location: "'Other'!A1", display: 'go' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const link = getHyperlink(ws2, 'C3');
    expect(link?.location).toBe("'Other'!A1");
    expect(link?.target).toBeUndefined();
    expect(link?.rId).toBeUndefined();
  });

  it('preserves a mix of external + internal links on one sheet', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setHyperlink(ws, 'A1', { target: 'https://one.example' });
    setHyperlink(ws, 'A2', { target: 'https://two.example' });
    setHyperlink(ws, 'A3', { location: 'Other!A1' });
    addWorksheet(wb, 'Other');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const refs = ws2.hyperlinks.map((h) => h.ref).sort();
    expect(refs).toEqual(['A1', 'A2', 'A3']);
    expect(getHyperlink(ws2, 'A1')?.target).toBe('https://one.example');
    expect(getHyperlink(ws2, 'A2')?.target).toBe('https://two.example');
    expect(getHyperlink(ws2, 'A3')?.location).toBe('Other!A1');
  });
});

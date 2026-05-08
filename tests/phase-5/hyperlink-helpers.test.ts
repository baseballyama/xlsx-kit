// Tests for the hyperlink builder helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  addInternalHyperlink,
  addMailtoHyperlink,
  addUrlHyperlink,
} from '../../src/xlsx/worksheet/hyperlinks';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addUrlHyperlink', () => {
  it('attaches an external URL with optional tooltip + display', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const hl = addUrlHyperlink(ws, 'A1', 'https://example.com', {
      tooltip: 'Open example',
      display: 'Open',
    });
    expect(hl.target).toBe('https://example.com');
    expect(hl.tooltip).toBe('Open example');
    expect(hl.display).toBe('Open');
    expect(ws.hyperlinks).toHaveLength(1);
  });

  it('replaces an existing hyperlink at the same ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'A1', 'https://example.com');
    addUrlHyperlink(ws, 'A1', 'https://other.example');
    expect(ws.hyperlinks).toHaveLength(1);
    expect(ws.hyperlinks[0]?.target).toBe('https://other.example');
  });
});

describe('addInternalHyperlink', () => {
  it('jumps to a sheet ref without a rels entry (location only)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const hl = addInternalHyperlink(ws, 'A1', "'Sheet 2'!A1", { tooltip: 'Jump' });
    expect(hl.location).toBe("'Sheet 2'!A1");
    expect(hl.target).toBeUndefined();
    expect(hl.tooltip).toBe('Jump');
  });

  it('replaces existing entries at the same ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'B2', 'https://x.test');
    addInternalHyperlink(ws, 'B2', "'Sheet 2'!Z9");
    expect(ws.hyperlinks).toHaveLength(1);
    expect(ws.hyperlinks[0]?.location).toBe("'Sheet 2'!Z9");
    expect(ws.hyperlinks[0]?.target).toBeUndefined();
  });
});

describe('addMailtoHyperlink', () => {
  it('builds a mailto: URL', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const hl = addMailtoHyperlink(ws, 'A1', 'qa@example.com');
    expect(hl.target).toBe('mailto:qa@example.com');
  });

  it('encodes a subject parameter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const hl = addMailtoHyperlink(ws, 'A1', 'qa@example.com', {
      subject: 'Bug report: failed import',
    });
    expect(hl.target).toBe('mailto:qa@example.com?subject=Bug%20report%3A%20failed%20import');
  });
});

describe('save → load round-trip', () => {
  it('preserves URL + internal + mailto entries', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'GitHub');
    setCell(ws, 2, 1, 'Sheet B');
    setCell(ws, 3, 1, 'Email');
    addUrlHyperlink(ws, 'A1', 'https://github.com');
    addInternalHyperlink(ws, 'A2', 'A!B5');
    addMailtoHyperlink(ws, 'A3', 'qa@example.com', { subject: 'Hi' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.hyperlinks.length).toBeGreaterThanOrEqual(2);
    const a1 = ws2.hyperlinks.find((h) => h.ref === 'A1');
    const a2 = ws2.hyperlinks.find((h) => h.ref === 'A2');
    const a3 = ws2.hyperlinks.find((h) => h.ref === 'A3');
    expect(a1?.target).toBe('https://github.com');
    expect(a2?.location).toBe('A!B5');
    expect(a3?.target).toBe('mailto:qa@example.com?subject=Hi');
  });
});
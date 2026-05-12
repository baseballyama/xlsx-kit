import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { makeChartsheet } from '../../src/chartsheet/chartsheet';
import { chartsheetToBytes, parseChartsheetXml, serializeChartsheet } from '../../src/chartsheet/chartsheet-xml';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addChartsheet, addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('Chartsheet XML round-trip', () => {
  it('preserves sheetView attributes', () => {
    const cs = makeChartsheet('Chart 1');
    cs.views = [{ workbookViewId: 0, tabSelected: true, zoomScale: 75, zoomToFit: false }];
    const bytes = chartsheetToBytes(cs);
    const back = parseChartsheetXml(bytes, 'Chart 1');
    expect(back.title).toBe('Chart 1');
    expect(back.views).toEqual(cs.views);
  });

  it('preserves sheetPr properties (published / codeName / tabColor)', () => {
    const cs = makeChartsheet('Tinted');
    cs.properties = { published: false, codeName: 'Sheet42', tabColor: { rgb: 'FF8800' } };
    const back = parseChartsheetXml(chartsheetToBytes(cs), 'Tinted');
    expect(back.properties).toEqual({ published: false, codeName: 'Sheet42', tabColor: { rgb: 'FF8800' } });
  });

  it('preserves sheetProtection (content + objects + algorithmName)', () => {
    const cs = makeChartsheet('Locked');
    cs.protection = {
      content: true,
      objects: true,
      algorithmName: 'SHA-512',
      hashValue: 'abc==',
      saltValue: 'def==',
      spinCount: 100000,
    };
    const back = parseChartsheetXml(chartsheetToBytes(cs), 'Locked');
    expect(back.protection).toEqual(cs.protection);
  });

  it('emits <drawing r:id> when caller passes drawingRId', () => {
    const cs = makeChartsheet('Visual');
    const xml = serializeChartsheet(cs, { drawingRId: 'rId7' });
    expect(xml).toContain('<drawing r:id="rId7"/>');
  });
});

describe('Chartsheet workbook integration', () => {
  it('emits xl/chartsheets/sheet1.xml + chartsheet rel + Override content type', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    const chart = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'Data!A1:A4' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
      title: 'Chartsheet Demo',
    });
    addChartsheet(wb, 'Chart 1', { chart: { space: chart } });

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);

    expect(entries['xl/worksheets/sheet1.xml']).toBeDefined();
    expect(entries['xl/chartsheets/sheet1.xml']).toBeDefined();
    expect(entries['xl/chartsheets/_rels/sheet1.xml.rels']).toBeDefined();
    expect(entries['xl/drawings/drawing1.xml']).toBeDefined();
    expect(entries['xl/charts/chart1.xml']).toBeDefined();

    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('chartsheet+xml');
    expect(ct).toContain('/xl/chartsheets/sheet1.xml');

    const wbRels = new TextDecoder().decode(entries['xl/_rels/workbook.xml.rels']);
    expect(wbRels).toContain('relationships/chartsheet');
    expect(wbRels).toContain('chartsheets/sheet1.xml');

    const csRels = new TextDecoder().decode(entries['xl/chartsheets/_rels/sheet1.xml.rels']);
    expect(csRels).toContain('relationships/drawing');
    expect(csRels).toContain('../drawings/drawing1.xml');
  });

  it('round-trips a workbook with worksheet + chartsheet through load', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    addChartsheet(wb, 'Chart 1', {
      chart: {
        space: makeChartSpace({
          plotArea: {
            chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'Data!A1:A2' } })] }),
            catAx: { axId: 1, crossAx: 2 },
            valAx: { axId: 2, crossAx: 1 },
          },
        }),
      },
    });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.length).toBe(2);
    expect(wb2.sheets[0]?.kind).toBe('worksheet');
    expect(wb2.sheets[1]?.kind).toBe('chartsheet');
    expect(wb2.sheets[1]?.sheet.title).toBe('Chart 1');

    const chartsheet = wb2.sheets[1];
    if (!chartsheet || chartsheet.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(chartsheet.sheet.drawing?.items.length).toBe(1);
    const item = chartsheet.sheet.drawing?.items[0];
    if (!item || item.content.kind !== 'chart') throw new Error('expected chart drawing item');
    expect(item.content.chart.space).toBeDefined();
  });

  it('addChartsheet rejects duplicate titles', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Same');
    expect(() => addChartsheet(wb, 'Same')).toThrowError(/already in use/);
  });
});

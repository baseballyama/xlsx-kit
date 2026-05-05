// Tests for visual conditional-formatting rule builders
// (colorScale / dataBar / iconSet).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  addColorScaleRule,
  addDataBarRule,
  addIconSetRule,
} from '../../src/worksheet/conditional-formatting';
import type { Worksheet } from '../../src/worksheet/worksheet';

const expectWorksheet = (
  s: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!s || !('rows' in s)) throw new Error('expected worksheet');
  return s as Worksheet;
};

describe('addColorScaleRule', () => {
  it('2-stop scale renders min + max cfvos and 2 colors', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addColorScaleRule(ws, 'A1:A10', {
      cfvos: [{ type: 'min' }, { type: 'max' }],
      colors: ['FFFF0000', 'FF00FF00'],
    });
    expect(rule.type).toBe('colorScale');
    expect(rule.priority).toBe(1);
    expect(rule.innerXml).toBe(
      '<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale>',
    );
  });

  it('3-stop scale with percentile midpoint', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addColorScaleRule(ws, 'A1:A10', {
      cfvos: [
        { type: 'min' },
        { type: 'percentile', val: '50' },
        { type: 'max' },
      ],
      colors: ['FFFF0000', 'FFFFFF00', 'FF00FF00'],
    });
    expect(rule.innerXml).toContain('<cfvo type="percentile" val="50"/>');
    expect(rule.innerXml).toContain('<color rgb="FFFFFF00"/>');
  });

  it('rejects mismatched cfvos / colors length', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() =>
      addColorScaleRule(ws, 'A1', {
        cfvos: [{ type: 'min' }, { type: 'max' }],
        colors: ['FFFF0000'],
      }),
    ).toThrow(/colors length.*must match/);
  });

  it('rejects 1-stop / 4-stop scales', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() =>
      addColorScaleRule(ws, 'A1', { cfvos: [{ type: 'min' }], colors: ['FFFF0000'] }),
    ).toThrow(/length 2 or 3/);
  });
});

describe('addDataBarRule', () => {
  it('default min/max with single color', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addDataBarRule(ws, 'B1:B5', { color: 'FF638EC6' });
    expect(rule.type).toBe('dataBar');
    expect(rule.innerXml).toBe(
      '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>',
    );
  });

  it('custom cfvos + minLength/maxLength/showValue attrs', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addDataBarRule(ws, 'B1:B5', {
      color: 'FFAA00FF',
      minCfvo: { type: 'percent', val: '10' },
      maxCfvo: { type: 'percent', val: '90' },
      minLength: 0,
      maxLength: 100,
      showValue: false,
    });
    expect(rule.innerXml).toContain('<dataBar minLength="0" maxLength="100" showValue="0">');
    expect(rule.innerXml).toContain('<cfvo type="percent" val="10"/>');
  });
});

describe('addIconSetRule', () => {
  it('3-icon traffic lights default', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addIconSetRule(ws, 'C1:C10', {
      iconSet: '3TrafficLights1',
      cfvos: [
        { type: 'percent', val: '0' },
        { type: 'percent', val: '33' },
        { type: 'percent', val: '67' },
      ],
    });
    expect(rule.type).toBe('iconSet');
    expect(rule.innerXml).toContain('iconSet="3TrafficLights1"');
    expect(rule.innerXml).toContain('<cfvo type="percent" val="33"/>');
  });

  it('5-arrow with reverse + showValue + percent attrs', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rule = addIconSetRule(ws, 'D1:D10', {
      iconSet: '5Arrows',
      cfvos: [
        { type: 'percent', val: '0' },
        { type: 'percent', val: '20' },
        { type: 'percent', val: '40' },
        { type: 'percent', val: '60' },
        { type: 'percent', val: '80' },
      ],
      reverse: true,
      showValue: false,
      percent: true,
    });
    expect(rule.innerXml).toContain(
      '<iconSet iconSet="5Arrows" reverse="1" showValue="0" percent="1">',
    );
  });

  it('rejects cfvos length out of [3, 5]', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() =>
      addIconSetRule(ws, 'A1', { iconSet: '3Arrows', cfvos: [{ type: 'percent', val: '50' }] }),
    ).toThrow(/length 3..5/);
    expect(() =>
      addIconSetRule(ws, 'A1', {
        iconSet: '3Arrows',
        cfvos: Array.from({ length: 6 }, () => ({ type: 'percent' as const, val: '50' })),
      }),
    ).toThrow(/length 3..5/);
  });
});

describe('visual CF rule round-trip through saveWorkbook → loadWorkbook', () => {
  it('colorScale + dataBar + iconSet survive a save/load cycle', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'CF');
    addColorScaleRule(ws, 'A1:A5', {
      cfvos: [{ type: 'min' }, { type: 'max' }],
      colors: ['FFFF0000', 'FF00FF00'],
    });
    addDataBarRule(ws, 'B1:B5', { color: 'FF638EC6' });
    addIconSetRule(ws, 'C1:C5', {
      iconSet: '3TrafficLights1',
      cfvos: [
        { type: 'percent', val: '0' },
        { type: 'percent', val: '33' },
        { type: 'percent', val: '67' },
      ],
    });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectWorksheet(wb2.sheets[0]?.sheet);
    const types = ws2.conditionalFormatting.flatMap((cf) => cf.rules.map((r) => r.type)).sort();
    expect(types).toEqual(['colorScale', 'dataBar', 'iconSet']);
    const rules = ws2.conditionalFormatting.flatMap((cf) => cf.rules);
    const colorRule = rules.find((r) => r.type === 'colorScale');
    expect(colorRule?.innerXml).toContain('FFFF0000');
    expect(colorRule?.innerXml).toContain('FF00FF00');
    const barRule = rules.find((r) => r.type === 'dataBar');
    expect(barRule?.innerXml).toContain('FF638EC6');
    const iconRule = rules.find((r) => r.type === 'iconSet');
    expect(iconRule?.innerXml).toContain('3TrafficLights1');
  });
});

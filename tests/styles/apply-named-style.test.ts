// Tests for applyBuiltinStyle / applyNamedStyle.

import { describe, expect, it } from 'vitest';
import {
  applyBuiltinStyle,
  applyNamedStyle,
  getCellFill,
  getCellFont,
  setCellFont,
} from '../../src/styles/cell-style';
import { addNamedStyle } from '../../src/styles/named-styles';
import { makeFont } from '../../src/styles/fonts';
import { makeColor } from '../../src/styles/colors';
import { makeFill, makePatternFill } from '../../src/styles/fills';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('applyBuiltinStyle', () => {
  it('"Good" applies the green palette + sets xfId', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'OK');
    applyBuiltinStyle(wb, c, 'Good');
    const xf = wb.styles.cellXfs[c.styleId];
    expect(xf?.xfId).toBeGreaterThanOrEqual(0);
    expect(getCellFont(wb, c).color?.rgb).toBe('FF006100');
    const fill = getCellFill(wb, c);
    expect(fill.kind === 'pattern' ? fill.fgColor?.rgb : undefined).toBe('FFC6EFCE');
  });

  it('"Bad" applies the red palette', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'NG');
    applyBuiltinStyle(wb, c, 'Bad');
    expect(getCellFont(wb, c).color?.rgb).toBe('FF9C0006');
  });

  it('idempotent: applying the same built-in twice reuses the same cellStyleXf entry', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 'a');
    const b = setCell(ws, 2, 1, 'b');
    applyBuiltinStyle(wb, a, 'Neutral');
    applyBuiltinStyle(wb, b, 'Neutral');
    expect(wb.styles.cellXfs[a.styleId]?.xfId).toBe(wb.styles.cellXfs[b.styleId]?.xfId);
    expect(wb.styles.namedStyles?.filter((s) => s.name === 'Neutral').length).toBe(1);
  });

  it('rejects unknown built-in name', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    expect(() => applyBuiltinStyle(wb, c, 'NotARealStyle')).toThrow(/unknown built-in style/);
  });
});

describe('applyNamedStyle', () => {
  it('finds a user-registered NamedStyle and applies its xf bundle', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'tag');
    addNamedStyle(wb.styles, {
      name: 'My Tag',
      font: makeFont({ bold: true, color: makeColor({ rgb: 'FFAA0000' }) }),
      fill: makeFill(makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFE0E0FF' }) })),
    });
    applyNamedStyle(wb, c, 'My Tag');
    expect(getCellFont(wb, c).bold).toBe(true);
    const fill = getCellFill(wb, c);
    expect(fill.kind === 'pattern' ? fill.fgColor?.rgb : undefined).toBe('FFE0E0FF');
  });

  it('throws when the named style is not registered', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    expect(() => applyNamedStyle(wb, c, 'Unknown')).toThrow(/no named style "Unknown"/);
  });

  it('a subsequent setCellFont composes onto the named-style baseline', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'mix');
    applyBuiltinStyle(wb, c, 'Good');
    setCellFont(wb, c, makeFont({ italic: true, color: makeColor({ rgb: 'FF006100' }) }));
    expect(getCellFont(wb, c).italic).toBe(true);
    // xfId should still point at the named style's cellStyleXfs entry.
    const xf = wb.styles.cellXfs[c.styleId];
    expect(xf?.xfId).toBeGreaterThanOrEqual(0);
  });
});

// Scenario 12: conditional formatting rules — cellIs, top10, color
// scale, data bar, icon set.
// Output: 12-conditional-format.xlsx
//
// What to verify in Excel:
// - Column A "Score" 1..20 with these rules applied:
//   - Cells > 15 → red background (cellIs greaterThan)
//   - Top 3 values → bold + green fill (top10 with rank=3)
//   - 3-color scale across the whole column.
// - Column B has a data bar showing relative magnitude.
// - Column C has the 5-arrows icon set (5 = ↑↑, 1 = ↓↓).

import { describe, expect, it } from 'vitest';
import { addCellXf, addDxf, addFill, addFont, defaultCellXf, makeColor, makeFont, makePatternFill } from '../../../src/xlsx/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { parseMultiCellRange } from '../../../src/xlsx/worksheet/cell-range';
import { writeWorkbook } from '../_helpers';

describe('e2e 12 — conditional formatting', () => {
  it('writes 12-conditional-format.xlsx', async () => {
    const wb = createWorkbook();
    addCellXf(wb.styles, defaultCellXf());
    const ws = addWorksheet(wb, 'CF');
    setCell(ws, 1, 1, 'Score');
    setCell(ws, 1, 2, 'Bar');
    setCell(ws, 1, 3, 'Icon');
    for (let r = 2; r <= 21; r++) {
      setCell(ws, r, 1, r - 1);
      setCell(ws, r, 2, r - 1);
      setCell(ws, r, 3, ((r - 2) % 5) + 1);
    }

    const redDxf = addDxf(wb.styles, {
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFC7CE' }) }),
      font: makeFont({ color: makeColor({ rgb: 'FF9C0006' }) }),
    });
    const greenDxf = addDxf(wb.styles, {
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFC6EFCE' }) }),
      font: makeFont({ bold: true, color: makeColor({ rgb: 'FF006100' }) }),
    });
    void addFill; void addFont; // silence unused-import lint when refactored

    ws.conditionalFormatting.push({
      sqref: parseMultiCellRange('A2:A21'),
      rules: [
        { type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['15'], dxfId: redDxf },
        { type: 'top10', priority: 2, rank: 3, formulas: [], dxfId: greenDxf },
        {
          type: 'colorScale',
          priority: 3,
          formulas: [],
          innerXml:
            '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
        },
      ],
    });

    ws.conditionalFormatting.push({
      sqref: parseMultiCellRange('B2:B21'),
      rules: [
        {
          type: 'dataBar',
          priority: 1,
          formulas: [],
          innerXml:
            '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>',
        },
      ],
    });

    ws.conditionalFormatting.push({
      sqref: parseMultiCellRange('C2:C21'),
      rules: [
        {
          type: 'iconSet',
          priority: 1,
          formulas: [],
          innerXml:
            '<iconSet iconSet="5Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="20"/><cfvo type="percent" val="40"/><cfvo type="percent" val="60"/><cfvo type="percent" val="80"/></iconSet>',
        },
      ],
    });

    const result = await writeWorkbook('12-conditional-format.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});

// Tests for getAllHyperlinks / getAllComments aggregators.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getAllComments,
  getAllHyperlinks,
} from '../../src/xlsx/workbook/workbook';
import { addUrlHyperlink } from '../../src/xlsx/worksheet/hyperlinks';
import { setComment } from '../../src/xlsx/worksheet/worksheet';

describe('getAllHyperlinks', () => {
  it('aggregates across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    addUrlHyperlink(a, 'A1', 'https://a-1.example');
    addUrlHyperlink(b, 'B2', 'https://b-2.example');
    addUrlHyperlink(a, 'C3', 'https://a-3.example');
    const out = getAllHyperlinks(wb).map(({ sheet, hyperlink }) => `${sheet.title}:${hyperlink.ref}`);
    expect(out).toEqual(['A:A1', 'A:C3', 'B:B2']);
  });

  it('skips chartsheets and empty worksheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A'); // empty
    addChartsheet(wb, 'Chart');
    expect(getAllHyperlinks(wb)).toEqual([]);
  });
});

describe('getAllComments', () => {
  it('aggregates across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setComment(a, { ref: 'A1', author: 'Alice', text: 'a1' });
    setComment(b, { ref: 'B1', author: 'Bob', text: 'b1' });
    setComment(a, { ref: 'A2', author: 'Alice', text: 'a2' });
    const out = getAllComments(wb).map(({ sheet, comment }) => `${sheet.title}:${comment.ref}`);
    expect(out).toEqual(['A:A1', 'A:A2', 'B:B1']);
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(getAllComments(wb)).toEqual([]);
  });
});

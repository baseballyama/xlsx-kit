// Apply font, fill, alignment, and a thin border to a header row.

import { saveWorkbook } from 'ooxml-js/xlsx/io';
import { toFile } from 'ooxml-js/node';
import {
  centerCell,
  setBold,
  setCellBackgroundColor,
  setCellBorderAll,
  setFontSize,
} from 'ooxml-js/xlsx/styles';
import { addWorksheet, createWorkbook } from 'ooxml-js/xlsx/workbook';
import { setCell } from 'ooxml-js/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const header = setCell(ws, 1, 1, 'Total revenue');
setBold(wb, header);
setFontSize(wb, header, 12);
setCellBackgroundColor(wb, header, 'FFE0E7FF');
centerCell(wb, header);
setCellBorderAll(wb, header, { style: 'thin' });

await saveWorkbook(wb, toFile('styled.xlsx'));

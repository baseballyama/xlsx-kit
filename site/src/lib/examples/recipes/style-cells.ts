// Apply font, fill, alignment, and a thin border to a header row.

import { saveWorkbook } from 'xlsxlite/io';
import { toFile } from 'xlsxlite/node';
import {
  centerCell,
  setBold,
  setCellBackgroundColor,
  setCellBorderAll,
  setFontSize,
} from 'xlsxlite/styles';
import { addWorksheet, createWorkbook } from 'xlsxlite/workbook';
import { setCell } from 'xlsxlite/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const header = setCell(ws, 1, 1, 'Total revenue');
setBold(wb, header);
setFontSize(wb, header, 12);
setCellBackgroundColor(wb, header, 'FFE0E7FF');
centerCell(wb, header);
setCellBorderAll(wb, header, { style: 'thin' });

await saveWorkbook(wb, toFile('styled.xlsx'));

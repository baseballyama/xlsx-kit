// Apply font, fill, alignment, and a thin border to a header row.

import { saveWorkbook } from 'xlsxify/io';
import { toFile } from 'xlsxify/node';
import {
  centerCell,
  setBold,
  setCellBackgroundColor,
  setCellBorderAll,
  setFontSize,
} from 'xlsxify/styles';
import { addWorksheet, createWorkbook } from 'xlsxify/workbook';
import { setCell } from 'xlsxify/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const header = setCell(ws, 1, 1, 'Total revenue');
setBold(wb, header);
setFontSize(wb, header, 12);
setCellBackgroundColor(wb, header, 'FFE0E7FF');
centerCell(wb, header);
setCellBorderAll(wb, header, { style: 'thin' });

await saveWorkbook(wb, toFile('styled.xlsx'));

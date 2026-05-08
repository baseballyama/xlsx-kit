// Apply font, fill, alignment, and a thin border to a header row.

import { saveWorkbook, toFile } from 'openxml-js/node';
import {
  centerCell,
  setBold,
  setCellBackgroundColor,
  setCellBorderAll,
  setFontSize,
} from 'openxml-js/styles';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { setCell } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const header = setCell(ws, 1, 1, 'Total revenue');
setBold(wb, header);
setFontSize(wb, header, 12);
setCellBackgroundColor(wb, header, 'FFE0E7FF');
centerCell(wb, header);
setCellBorderAll(wb, header, { style: 'thin' });

await saveWorkbook(wb, toFile('styled.xlsx'));

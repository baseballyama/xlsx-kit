// Apply number formats: currency, percentage, and a date-time.

import { saveWorkbook } from 'ooxml-js/xlsx/io';
import { toFile } from 'ooxml-js/node';
import {
  FORMAT_DATE_DATETIME,
  setCellAsCurrency,
  setCellAsPercent,
  setCellNumberFormat,
} from 'ooxml-js/xlsx/styles';
import { addWorksheet, createWorkbook } from 'ooxml-js/xlsx/workbook';
import { setCell } from 'ooxml-js/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Numbers');

setCellAsCurrency(wb, setCell(ws, 1, 1, 12_400), { symbol: '$' });
setCellAsPercent(wb, setCell(ws, 1, 2, 0.187), 1);

const dateCell = setCell(ws, 1, 3, new Date('2026-05-08T09:30:00Z'));
setCellNumberFormat(wb, dateCell, FORMAT_DATE_DATETIME);

await saveWorkbook(wb, toFile('numbers.xlsx'));

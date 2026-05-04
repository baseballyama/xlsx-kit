import { describe, expect, it } from 'vitest';
import {
  dateToExcel,
  durationToExcel,
  excelToDate,
  excelToDuration,
  fromIso8601,
  MAC_EPOCH_MS,
  toIso8601,
  WINDOWS_EPOCH_MS,
} from '../../../src/utils/datetime';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';

const utcDate = (y: number, m: number, d: number, h = 0, min = 0, s = 0): Date =>
  new Date(Date.UTC(y, m - 1, d, h, min, s));

describe('excelToDate / dateToExcel — Windows 1900 epoch', () => {
  // Cross-checked against openpyxl's from_excel / to_excel on
  // representative serials including the leap-bug boundary.
  const cases: ReadonlyArray<readonly [number, Date]> = [
    [1, utcDate(1900, 1, 1)],
    [2, utcDate(1900, 1, 2)],
    [59, utcDate(1900, 2, 28)],
    // openpyxl collapses serial 60 (phantom 1900-02-29) onto 1900-02-28.
    [60, utcDate(1900, 2, 28)],
    [61, utcDate(1900, 3, 1)],
    [367, utcDate(1901, 1, 1)],
    [40000, utcDate(2009, 7, 6)],
    [44927, utcDate(2023, 1, 1)],
  ];

  it.each(cases)('excelToDate(%i) === %s', (serial, expected) => {
    const got = excelToDate(serial);
    expect(got.toISOString()).toBe(expected.toISOString());
  });

  it.each(cases.filter(([s]) => s !== 60))('dateToExcel inverts excelToDate (%i)', (serial, _date) => {
    const back = dateToExcel(excelToDate(serial));
    expect(back).toBeCloseTo(serial, 9);
  });

  it('dateToExcel for Feb 28, 1900 returns serial 59 (phantom collapse)', () => {
    expect(dateToExcel(utcDate(1900, 2, 28))).toBe(59);
  });

  it('handles fractional time-of-day (0.5 = noon)', () => {
    const d = excelToDate(40000.5);
    expect(d.getUTCHours()).toBe(12);
    expect(d.getUTCMinutes()).toBe(0);
  });

  it('round-trips a datetime with seconds', () => {
    const d = utcDate(2023, 6, 15, 13, 45, 30);
    const serial = dateToExcel(d);
    const back = excelToDate(serial);
    expect(back.toISOString()).toBe(d.toISOString());
  });
});

describe('excelToDate / dateToExcel — Mac 1904 epoch', () => {
  // Mac 1904: serial 0 == 1904-01-01.
  it('serial 0 maps to 1904-01-01', () => {
    expect(excelToDate(0, { epoch: 'mac' }).toISOString()).toBe(utcDate(1904, 1, 1).toISOString());
  });

  it('serial 1 maps to 1904-01-02 (no leap bug)', () => {
    expect(excelToDate(1, { epoch: 'mac' }).toISOString()).toBe(utcDate(1904, 1, 2).toISOString());
  });

  it('round-trips via dateToExcel under the mac epoch', () => {
    const d = utcDate(2026, 5, 4, 10, 0, 0);
    const serial = dateToExcel(d, { epoch: 'mac' });
    const back = excelToDate(serial, { epoch: 'mac' });
    expect(back.toISOString()).toBe(d.toISOString());
  });

  it('windows and mac serials for the same date differ by 1462', () => {
    const d = utcDate(2023, 1, 1);
    const winSerial = dateToExcel(d);
    const macSerial = dateToExcel(d, { epoch: 'mac' });
    expect(winSerial - macSerial).toBe(1462);
  });
});

describe('exceptional inputs', () => {
  it('excelToDate throws on non-finite inputs', () => {
    expect(() => excelToDate(Number.NaN)).toThrowError(OpenXmlSchemaError);
    expect(() => excelToDate(Number.POSITIVE_INFINITY)).toThrowError(OpenXmlSchemaError);
  });

  it('dateToExcel throws on an invalid Date', () => {
    expect(() => dateToExcel(new Date(Number.NaN))).toThrowError(OpenXmlSchemaError);
  });
});

describe('excelToDuration / durationToExcel', () => {
  it('1.0 day == 86_400_000 ms', () => {
    expect(excelToDuration(1)).toBe(86_400_000);
    expect(durationToExcel(86_400_000)).toBe(1);
  });

  it('0.5 day == 43_200_000 ms (12 hours)', () => {
    expect(excelToDuration(0.5)).toBe(43_200_000);
  });

  it('round-trips through ms', () => {
    const ms = 12_345_678;
    expect(excelToDuration(durationToExcel(ms))).toBe(ms);
  });
});

describe('ISO 8601 helpers', () => {
  it('parses "2010-07-28T08:40:37Z"', () => {
    const d = fromIso8601('2010-07-28T08:40:37Z');
    expect(d.getUTCFullYear()).toBe(2010);
    expect(d.getUTCMonth()).toBe(6);
    expect(d.getUTCDate()).toBe(28);
    expect(d.getUTCHours()).toBe(8);
    expect(d.getUTCMinutes()).toBe(40);
    expect(d.getUTCSeconds()).toBe(37);
  });

  it('toIso8601 trims the millisecond suffix', () => {
    const d = utcDate(2010, 7, 28, 8, 40, 37);
    expect(toIso8601(d)).toBe('2010-07-28T08:40:37Z');
  });

  it('round-trips through fromIso8601', () => {
    const s = '2026-05-04T13:30:00Z';
    expect(toIso8601(fromIso8601(s))).toBe(s);
  });

  it('rejects non-strings / empty / unparseable input', () => {
    expect(() => fromIso8601('')).toThrowError(OpenXmlSchemaError);
    expect(() => fromIso8601('not a date')).toThrowError(OpenXmlSchemaError);
    expect(() => toIso8601(new Date(Number.NaN))).toThrowError(OpenXmlSchemaError);
  });
});

describe('epoch constants', () => {
  it('WINDOWS_EPOCH_MS / MAC_EPOCH_MS are correct', () => {
    expect(new Date(WINDOWS_EPOCH_MS).toISOString()).toBe('1899-12-30T00:00:00.000Z');
    expect(new Date(MAC_EPOCH_MS).toISOString()).toBe('1904-01-01T00:00:00.000Z');
  });
});

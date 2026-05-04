// Excel <-> JavaScript Date conversions. Mirrors
// openpyxl/openpyxl/utils/datetime.py.
//
// Excel stores datetimes as fractional "serial days" since an epoch.
// Two epochs are in use:
//   * Windows  ("1900 date system", default): epoch 1899-12-30 with the
//     well-known 1900 leap-year bug — Excel treats 1900-02-29 as a
//     valid day even though it isn't. We collapse that phantom day
//     onto 1900-02-28 (the openpyxl approach), so date math past
//     March 1, 1900 stays consistent without leaking the bug.
//   * Mac     ("1904 date system"): epoch 1904-01-01, no leap bug.
//
// Per docs/plan/03-foundations.md §7.2 dates stay as numeric serials
// on the worksheet hot path; conversion to a JS Date happens lazily
// only when callers ask. JS Dates are interpreted in UTC throughout
// to avoid timezone drift between read and write.

import { OpenXmlSchemaError } from './exceptions';

/** Excel epoch identifier. */
export type ExcelEpoch = 'windows' | 'mac';

/** 1899-12-30 (UTC) — the Windows / 1900-system epoch in ms. */
export const WINDOWS_EPOCH_MS = Date.UTC(1899, 11, 30);
/** 1904-01-01 (UTC) — the Mac / 1904-system epoch in ms. */
export const MAC_EPOCH_MS = Date.UTC(1904, 0, 1);

const MS_PER_DAY = 86_400_000;
/** Serial day index of the phantom 1900-02-29; absorbed onto 1900-02-28. */
const LEAP_DUPLICATE_DAY = 60;

const epochMs = (e: ExcelEpoch | undefined): number => (e === 'mac' ? MAC_EPOCH_MS : WINDOWS_EPOCH_MS);

/**
 * Convert an Excel serial date into a JS `Date` (UTC). The fractional
 * part is treated as a fraction of a day. For Windows 1900 the leap-bug
 * compensation kicks in for serials in [0, 60).
 */
export function excelToDate(serial: number, opts?: { epoch?: ExcelEpoch }): Date {
  if (!Number.isFinite(serial)) {
    throw new OpenXmlSchemaError(`excelToDate: serial "${serial}" is not finite`);
  }
  const epoch = epochMs(opts?.epoch);
  const day = Math.floor(serial);
  const fraction = serial - day;
  const isWindows = epoch === WINDOWS_EPOCH_MS;
  // Windows quirk: bump days [0, 60) by one so the day count lines up
  // with Excel's serial numbering through the phantom Feb 29 1900.
  const dayAdjusted = isWindows && serial >= 0 && serial < LEAP_DUPLICATE_DAY ? day + 1 : day;
  const ms = epoch + dayAdjusted * MS_PER_DAY + Math.round(fraction * MS_PER_DAY);
  return new Date(ms);
}

/**
 * Convert a JS `Date` into an Excel serial. The Date is read in UTC.
 * On Windows 1900, dates ≤ 1900-02-28 get a -1 day correction to
 * account for Excel's phantom leap day.
 */
export function dateToExcel(date: Date, opts?: { epoch?: ExcelEpoch }): number {
  const t = date.getTime();
  if (!Number.isFinite(t)) {
    throw new OpenXmlSchemaError('dateToExcel: invalid Date');
  }
  const epoch = epochMs(opts?.epoch);
  const dayStartMs = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  const days = Math.round((dayStartMs - epoch) / MS_PER_DAY);
  const subDay = (t - dayStartMs) / MS_PER_DAY;
  const isWindows = epoch === WINDOWS_EPOCH_MS;
  const daysAdjusted = isWindows && days <= LEAP_DUPLICATE_DAY ? days - 1 : days;
  return daysAdjusted + subDay;
}

/** Excel duration serial (fraction of a day) → milliseconds. */
export function excelToDuration(serial: number): number {
  if (!Number.isFinite(serial)) {
    throw new OpenXmlSchemaError(`excelToDuration: serial "${serial}" is not finite`);
  }
  return Math.round(serial * MS_PER_DAY);
}

/** Milliseconds → Excel duration serial (fraction of a day). */
export function durationToExcel(ms: number): number {
  if (!Number.isFinite(ms)) {
    throw new OpenXmlSchemaError(`durationToExcel: ms "${ms}" is not finite`);
  }
  return ms / MS_PER_DAY;
}

// ---- ISO 8601 helpers ------------------------------------------------------

/**
 * Parse an ISO-8601 / W3CDTF datetime string into a `Date`. Same
 * grammar as `new Date(string)`; the wrapper just adds typed error
 * reporting and a stricter "must be a recognised ISO" guard.
 */
export function fromIso8601(s: string): Date {
  if (typeof s !== 'string' || s.length === 0) {
    throw new OpenXmlSchemaError(`fromIso8601: empty input`);
  }
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) {
    throw new OpenXmlSchemaError(`fromIso8601: invalid datetime "${s}"`);
  }
  return d;
}

/**
 * Format a `Date` as ISO-8601 with second precision in UTC. Trims the
 * millisecond fragment that `Date.toISOString()` always produces, so
 * the output matches Excel / openpyxl's W3CDTF style.
 */
export function toIso8601(d: Date): string {
  if (Number.isNaN(d.getTime())) {
    throw new OpenXmlSchemaError('toIso8601: invalid Date');
  }
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

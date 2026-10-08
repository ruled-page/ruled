import { displayFormula } from '../core/formula';
import { parseInput } from '../core/input';
import { format, formatColor, isDateFormat } from 'numfmt';
import { VT, type DisplayStyle } from '../core/types';

export interface CellRecord {
  t: number;
  n?: number;
  s?: string;
  style: number;
  /** Formula in file form. */
  f?: string;
  flags: number;
}

const dateFormats = new Map<string, boolean>();
function isDate(pattern: string): boolean {
  let v = dateFormats.get(pattern);
  if (v === undefined) {
    try {
      v = isDateFormat(pattern);
    } catch {
      v = false;
    }
    dateFormats.set(pattern, v);
  }
  return v;
}

export function isDateStyle(style: DisplayStyle): boolean {
  return isDate(style.numFmt);
}

const OPTS = { throws: false, leap1900: true, nbsp: false } as const;

const tidy = (n: number) => parseFloat(n.toPrecision(15));
const isPercent = (pattern: string) => /%/.test(pattern.replace(/"[^"]*"/g, ''));

/**
 * The value with every stored digit (to Excel's 15 significant digits),
 * keeping percentages as percentages. Dates and text are unaffected.
 */
export function exactText(rec: CellRecord, style: DisplayStyle, grouping: boolean): string | undefined {
  if (rec.t !== VT.Number || isDate(style.numFmt)) return undefined;
  const n = rec.n ?? 0;
  if (isPercent(style.numFmt)) return `${tidy(n * 100)}%`;
  return grouping ? tidy(n).toLocaleString(undefined, { maximumFractionDigits: 15 }) : String(tidy(n));
}

/** The text Excel would show for a cell, and the colour its number format asks for. */
export function displayText(rec: CellRecord, style: DisplayStyle, date1904: boolean, exact = false): { text: string; color?: string } {
  if (exact) {
    const text = exactText(rec, style, true);
    if (text !== undefined) return { text };
  }
  switch (rec.t) {
    case VT.Number: {
      const pattern = style.numFmt || 'General';
      // Formats work in the 1900 date system; shift 1904-system serials into it.
      const n = date1904 && isDate(pattern) ? (rec.n ?? 0) + 1462 : (rec.n ?? 0);
      let text: string;
      try {
        text = format(pattern, n, OPTS);
      } catch {
        text = String(n);
      }
      let color: string | undefined;
      try {
        color = (formatColor(pattern, n, { throws: false }) as string | null) ?? undefined;
      } catch {
        color = undefined;
      }
      return { text, color };
    }
    case VT.String: {
      const s = rec.s ?? '';
      if (style.numFmt && style.numFmt !== 'General' && style.numFmt.includes('@')) {
        try {
          return { text: format(style.numFmt, s, OPTS) };
        } catch {
          return { text: s };
        }
      }
      return { text: s };
    }
    case VT.Bool:
      return { text: rec.n === 1 ? 'TRUE' : 'FALSE' };
    case VT.Error:
      return { text: rec.s ?? '#VALUE!' };
    default:
      return { text: '' };
  }
}

/** Excel's default alignment: numbers right, booleans and errors centred, text left. */
export function alignment(rec: CellRecord, style: DisplayStyle): 'left' | 'right' | 'center' {
  switch (style.hAlign) {
    case 'center':
    case 'centerContinuous':
    case 'distributed':
      return 'center';
    case 'right':
      return 'right';
    case 'left':
    case 'fill':
    case 'justify':
      return 'left';
  }
  if (rec.t === VT.Number) return 'right';
  if (rec.t === VT.Bool || rec.t === VT.Error) return 'center';
  return 'left';
}

/**
 * A date, with its time if it has one, as the user's locale writes it, so
 * that it reads back as the same date when the edit is committed: 03/01/2026
 * in the US, 01.03.2026 in Germany. ISO (2026-03-01) where the locale's own
 * form wouldn't read back.
 */
function editableDate(serial: number, date1904: boolean, locale: string): string {
  const seconds = Math.round(serial * 86_400);
  const at = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) + Math.floor(seconds / 86_400) * 86_400_000);
  const s = seconds % 86_400;
  const pad = (n: number) => String(n).padStart(2, '0');
  const time = s ? ` ${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}${s % 60 ? `:${pad(s % 60)}` : ''}` : '';
  const iso = `${at.getUTCFullYear()}-${pad(at.getUTCMonth() + 1)}-${pad(at.getUTCDate())}${time}`;
  let local: string;
  try {
    local = new Intl.DateTimeFormat(locale, { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' }).format(at) + time;
  } catch {
    return iso;
  }
  const back = parseInput(local, 'General', locale, date1904).content;
  return back.kind === 'number' && Math.abs(back.value - seconds / 86_400) < 1e-6 ? local : iso;
}

/**
 * What the cell editor and formula bar show: the formula, or the value as
 * typed. Dates are written the way the user's locale reads them, whatever
 * the cell's own format, so that committing an unchanged date keeps it.
 */
export function editText(rec: CellRecord | null, style: DisplayStyle, date1904: boolean, locale = 'en-US'): string {
  if (!rec) return '';
  if (rec.f !== undefined) return displayFormula(rec.f);
  if (rec.t === VT.Number) {
    const n = rec.n ?? 0;
    // Before 1 March 1900 the 1900 system has Excel's phantom 29 February; times alone have no date.
    if (isDate(style.numFmt)) return n >= 1 && (date1904 || n >= 61) ? editableDate(n, date1904, locale) : displayText(rec, style, date1904).text;
    if (isPercent(style.numFmt)) return `${tidy((rec.n ?? 0) * 100)}%`;
    return String(rec.n ?? '');
  }
  if (rec.t === VT.String) {
    const s = rec.s ?? '';
    // A leading apostrophe keeps text that looks like a number or formula as text.
    return /^[=+\-@]|^\s*[\d.,$%]+\s*$|^(true|false)$/i.test(s) ? "'" + s : s;
  }
  return displayText(rec, style, date1904).text;
}

/**
 * A cell's value as JSON: numbers as numbers unless formatted as dates (then
 * ISO text, with the time when there is one), logicals as booleans, errors
 * as their text, empty cells as null.
 */
export function cellJson(rec: CellRecord | null, style: DisplayStyle, date1904: boolean): string | number | boolean | null {
  if (!rec) return null;
  switch (rec.t) {
    case VT.Number: {
      const n = rec.n ?? 0;
      if (!isDate(style.numFmt)) return n;
      const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
      const iso = new Date(base + Math.round(n * 86_400_000)).toISOString();
      return Number.isInteger(n) ? iso.slice(0, 10) : iso.replace(/\.000Z$/, 'Z');
    }
    case VT.String:
    case VT.Error:
      return rec.s ?? '';
    case VT.Bool:
      return rec.n === 1;
    default:
      return null;
  }
}

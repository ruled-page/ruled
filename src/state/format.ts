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

/** What the cell editor and formula bar show: the formula, or the value as typed. */
export function editText(rec: CellRecord | null, style: DisplayStyle, date1904: boolean): string {
  if (!rec) return '';
  if (rec.f !== undefined) return '=' + rec.f.replace(/_xlfn\.(?:_xlws\.)?|_xlws\./gi, '');
  if (rec.t === VT.Number) {
    if (isDate(style.numFmt)) return displayText(rec, style, date1904).text;
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

/** How a locale writes dates and decimals, for reading typed and pasted text. */
export interface LocaleInfo {
  /** 3/4/2026 means 3 April (true) or 4 March (false). */
  dayFirst: boolean;
  /** 1.234,5 rather than 1,234.5. */
  decimalComma: boolean;
}

const cache = new Map<string, LocaleInfo>();

export function localeInfo(locale: string | undefined): LocaleInfo {
  const key = locale || 'en-US';
  let info = cache.get(key);
  if (info) return info;
  try {
    const parts = new Intl.DateTimeFormat(key, { day: 'numeric', month: 'numeric', year: 'numeric' }).formatToParts(new Date(2001, 10, 22));
    const day = parts.findIndex((p) => p.type === 'day');
    const month = parts.findIndex((p) => p.type === 'month');
    const decimal = new Intl.NumberFormat(key).formatToParts(1.5).find((p) => p.type === 'decimal')?.value;
    info = { dayFirst: day < month, decimalComma: decimal === ',' };
  } catch {
    info = { dayFirst: false, decimalComma: false };
  }
  cache.set(key, info);
  return info;
}

const SLASH_DATE = /^\s*(\d{1,2})([/.-])(\d{1,2})\2(\d{2}|\d{4})(\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)?\s*$/;

function expandYear(y: string): number {
  if (y.length === 4) return Number(y);
  const n = Number(y);
  // Excel's rule for two-digit years: 00–29 are 2000s, 30–99 are 1900s.
  return n < 30 ? 2000 + n : 1900 + n;
}

export interface LocalDate {
  year: number;
  month: number;
  day: number;
  time?: string;
  /** Both readings are valid dates and differ, e.g. 03/04/2026. */
  ambiguous: boolean;
}

/** Reads d/m/y or m/d/y text in the locale's order. */
export function readLocalDate(text: string, info: LocaleInfo): LocalDate | undefined {
  const m = SLASH_DATE.exec(text);
  if (!m) return undefined;
  const a = Number(m[1]);
  const b = Number(m[3]);
  const [day, month] = info.dayFirst ? [a, b] : [b, a];
  const year = expandYear(m[4]);
  const valid = (d: number, mo: number) => mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(Date.UTC(year, mo, 0)).getUTCDate();
  if (!valid(day, month)) return undefined;
  return { year, month, day, time: m[5]?.trim(), ambiguous: a !== b && valid(b, a) && a <= 12 && b <= 12 };
}

/**
 * Rewrites text typed in the user's locale into the form the parser
 * understands: dates to ISO, and decimal-comma numbers to decimal points.
 * Returns the number format the cell should get, if any.
 */
export function normalizeForLocale(text: string, info: LocaleInfo): { text: string; numFmt?: string } {
  const date = readLocalDate(text, info);
  if (date) {
    const iso = `${date.year}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`;
    const base = info.dayFirst ? 'dd/mm/yyyy' : 'm/d/yyyy';
    return { text: date.time ? `${iso} ${date.time}` : iso, numFmt: date.time ? `${base} h:mm` : base };
  }
  if (info.decimalComma) {
    const m = /^\s*([-+]?)([\d.\s  ]*\d)(?:,(\d+))?(\s*%)?\s*$/.exec(text);
    if (m && (m[3] !== undefined || /\.\d{3}(?!\d)/.test(m[2]))) {
      const whole = m[2].replace(/[.\s  ]/g, '');
      return { text: `${m[1]}${whole}${m[3] !== undefined ? '.' + m[3] : ''}${m[4] ? '%' : ''}` };
    }
  }
  return { text };
}

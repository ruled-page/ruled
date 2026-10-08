import { builtinNumFmt } from '../core/styles';

/** The number formats the format bar offers, as Excel's Home tab lists them. */
export type NumberFormatKind = 'general' | 'number' | 'currency' | 'percent' | 'shortDate' | 'longDate' | 'time' | 'scientific' | 'text';

export const NUMBER_FORMAT_KINDS: NumberFormatKind[] = ['general', 'number', 'currency', 'percent', 'shortDate', 'longDate', 'time', 'scientific', 'text'];

/** Currency formats by region, as Excel writes them whatever the display language. */
const CURRENCIES: Record<string, string> = {
  US: '"$"#,##0.00',
  CA: '"$"#,##0.00',
  AU: '"$"#,##0.00',
  NZ: '"$"#,##0.00',
  GB: '"£"#,##0.00',
  IE: '"€"#,##0.00',
  CH: '"CHF" #,##0.00',
  JP: '"¥"#,##0',
  CN: '"¥"#,##0.00',
  IN: '"₹"#,##0.00',
  SE: '#,##0.00 "kr"',
  NO: '#,##0.00 "kr"',
  DK: '#,##0.00 "kr."',
  PL: '#,##0.00 "zł"',
};
const EURO = new Set(['AT', 'BE', 'DE', 'EE', 'ES', 'FI', 'FR', 'GR', 'HR', 'IT', 'LT', 'LU', 'LV', 'MT', 'NL', 'PT', 'SI', 'SK']);

function region(locale: string): string {
  try {
    return new Intl.Locale(locale).maximize().region ?? 'US';
  } catch {
    return 'US';
  }
}

/** A kind's format code for the user's locale. */
export function numberFormatCode(kind: NumberFormatKind, locale: string): string {
  switch (kind) {
    case 'general':
      return 'General';
    case 'number':
      return '#,##0.00';
    case 'currency': {
      const r = region(locale);
      return CURRENCIES[r] ?? (EURO.has(r) ? '#,##0.00 "€"' : '"$"#,##0.00');
    }
    case 'percent':
      return '0.00%';
    case 'shortDate':
      return builtinNumFmt(14, locale);
    case 'longDate':
      return 'dddd, mmmm d, yyyy';
    case 'time':
      return 'h:mm:ss AM/PM';
    case 'scientific':
      return '0.00E+00';
    case 'text':
      return '@';
  }
}

/** Which of the kinds a format code is, if any. */
export function numberFormatKind(code: string, locale: string): NumberFormatKind | undefined {
  return NUMBER_FORMAT_KINDS.find((k) => numberFormatCode(k, locale) === code);
}

/**
 * The format code with one more (or one fewer) decimal place, as Excel's
 * Increase and Decrease Decimal buttons do, in every section of the code.
 * Formats without a number in them, such as dates and text, stay as they are.
 */
export function stepDecimals(code: string, delta: 1 | -1): string {
  if (code === 'General') return delta > 0 ? '0.0' : '0';
  return code
    .split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/)
    .map((section) => stepSection(section, delta))
    .join(';');
}

function stepSection(section: string, delta: 1 | -1): string {
  // Where the number's digits are, outside quoted text, [colours] and escaped characters.
  let last = -1;
  let point = -1;
  let date = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      i = section.indexOf('"', i + 1);
      if (i < 0) break;
    } else if (ch === '[') {
      i = section.indexOf(']', i + 1);
      if (i < 0) break;
    } else if (ch === '\\') i++;
    // The exponent's digits are not the number's.
    else if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-')) break;
    else if (ch === '0' || ch === '#' || ch === '?') last = i;
    else if (ch === '.' && point < 0) point = i;
    else if (/[dmyhsDMYHS@]/.test(ch)) date = true;
  }
  if (date || last < 0) return section;
  if (point < 0 || point > last) {
    return delta > 0 ? section.slice(0, last + 1) + '.0' + section.slice(last + 1) : section;
  }
  if (delta > 0) return section.slice(0, last + 1) + '0' + section.slice(last + 1);
  // One fewer: the last decimal goes, and the point with it when it was the only one.
  return last === point + 1 ? section.slice(0, point) + section.slice(last + 1) : section.slice(0, last) + section.slice(last + 1);
}

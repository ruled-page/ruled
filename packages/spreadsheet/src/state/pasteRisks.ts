import { localeInfo, normalizeForLocale, readLocalDate } from '../core/locale';
import { msg, type Text } from '../i18n/message';

/** A pasted value that would not be stored exactly as it reads. */
export interface PasteRisk {
  text: string;
  /** How the editor would store it. */
  becomes: string;
  /** Why: an ambiguous date, too many digits, or leading zeros. */
  reason: Text;
}

/**
 * Values that pasting would silently change: dates whose day and month could
 * be swapped, numbers longer than Excel's 15 significant digits, and
 * numbers whose leading zeros (often codes or IDs) would be dropped.
 * Dates are read in `locale` and shown, with the month named, in `language`.
 */
export function pasteRisks(values: readonly (readonly string[])[], locale: string, language = locale): PasteRisk[] {
  const info = localeInfo(locale);
  const dates = new Intl.DateTimeFormat(language || undefined, { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' });
  const out: PasteRisk[] = [];
  const seen = new Set<string>();
  for (const line of values) {
    for (const raw of line) {
      const text = raw.trim();
      if (!text || text.startsWith('=') || text.startsWith("'") || seen.has(text)) continue;
      seen.add(text);
      const date = readLocalDate(text, info);
      if (date?.ambiguous) {
        out.push({ text, becomes: dates.format(Date.UTC(date.year, date.month - 1, date.day)), reason: msg('store.paste.ambiguousDate') });
        continue;
      }
      // The number as the editor reads it in this locale: 1.234,5 is 1234.5 in German.
      const number = normalizeForLocale(text, info).text;
      const digits = /^[-+]?(\d[\d,]*)(\.\d+)?$/.exec(number);
      if (!digits) continue;
      const significant = (digits[1] + (digits[2] ?? '')).replace(/[,.]/g, '').replace(/^0+/, '').replace(/0+$/, '');
      const becomes = String(Number(number.replace(/,/g, '')));
      if (significant.length > 15) {
        out.push({ text, becomes, reason: msg('store.paste.tooManyDigits') });
      } else if (/^0\d/.test(digits[1]) && !digits[2]) {
        out.push({ text, becomes, reason: msg('store.paste.leadingZeros') });
      }
    }
  }
  return out;
}

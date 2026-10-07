import { localeInfo, readLocalDate } from '../core/locale';

/** A pasted value that would not be stored exactly as it reads. */
export interface PasteRisk {
  text: string;
  /** How the editor would store it. */
  becomes: string;
  reason: 'ambiguous date' | 'too many digits' | 'leading zeros';
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Values that pasting would silently change: dates whose day and month could
 * be swapped, numbers longer than Excel's 15 significant digits, and
 * numbers whose leading zeros (often codes or IDs) would be dropped.
 */
export function pasteRisks(values: readonly (readonly string[])[], locale: string): PasteRisk[] {
  const info = localeInfo(locale);
  const out: PasteRisk[] = [];
  const seen = new Set<string>();
  for (const line of values) {
    for (const raw of line) {
      const text = raw.trim();
      if (!text || text.startsWith('=') || text.startsWith("'") || seen.has(text)) continue;
      seen.add(text);
      const date = readLocalDate(text, info);
      if (date?.ambiguous) {
        out.push({ text, becomes: `${date.day} ${MONTHS[date.month - 1]} ${date.year}`, reason: 'ambiguous date' });
        continue;
      }
      const digits = /^[-+]?(\d[\d,]*)(\.\d+)?$/.exec(text);
      if (!digits) continue;
      const significant = (digits[1] + (digits[2] ?? '')).replace(/[,.]/g, '').replace(/^0+/, '').replace(/0+$/, '');
      if (significant.length > 15) {
        out.push({ text, becomes: String(Number(text.replace(/,/g, ''))), reason: 'too many digits' });
      } else if (/^0\d/.test(digits[1]) && !digits[2]) {
        out.push({ text, becomes: String(Number(text)), reason: 'leading zeros' });
      }
    }
  }
  return out;
}

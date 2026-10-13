import { ENGLISH_LISTS } from '../core/fill';
import { EXCEL_SHORT_NAMES } from '../core/locale';

const named = new Map<string, string[][]>();

/**
 * Month and day names, short and full, as the language writes them, from
 * Intl: Januar and Jan., Montag and Mo. in German. Each is taken as it
 * stands alone and as a date writes it, which differ in some languages
 * ("Jan" and "Jan.", or another grammatical case), and short names written
 * with a full stop come without it too, as they are often typed that way.
 * Excel's own short names follow, where they differ ("Mrz").
 */
export function namesIn(language: string): string[][] {
  let lists = named.get(language);
  if (lists) return lists;
  try {
    const names = (options: Intl.DateTimeFormatOptions, type: 'month' | 'weekday', times: number[]) => {
      const f = new Intl.DateTimeFormat(language, { ...options, timeZone: 'UTC' });
      return times.map((time) => f.formatToParts(time).find((p) => p.type === type)?.value ?? '');
    };
    const months = Array.from({ length: 12 }, (_, i) => Date.UTC(2023, i, 15));
    // 1 January 2023 was a Sunday, where Excel's lists of days start.
    const days = Array.from({ length: 7 }, (_, i) => Date.UTC(2023, 0, 1 + i));
    const widths = ['short', 'long'] as const;
    lists = distinct(
      [
        ...widths.flatMap((month) => forms(names({ month }, 'month', months), names({ month, day: 'numeric' }, 'month', months))),
        ...widths.flatMap((weekday) => forms(names({ weekday }, 'weekday', days), names({ weekday, day: 'numeric' }, 'weekday', days))),
        ...Object.values(EXCEL_SHORT_NAMES[language.split('-')[0].toLowerCase()] ?? {}),
      ].filter((list) => list.every(Boolean)),
    );
  } catch {
    lists = [];
  }
  named.set(language, lists);
  return lists;
}

/** A list of names as they stand alone, then as a date writes them, without and with their full stops. */
function forms(alone: string[], inDate: string[]): string[][] {
  return [alone, inDate.map((w) => w.replace(/\.$/, '')), inDate];
}

/** Excel's built-in lists in each language given, then in English, each list once. */
export function builtInLists(languages: readonly string[]): string[][] {
  return distinct([...languages.flatMap(namesIn), ...ENGLISH_LISTS.map((l) => [...l])]);
}

function distinct(lists: string[][]): string[][] {
  const seen = new Set<string>();
  return lists.filter((list) => {
    const key = list.join('\n').toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/**
 * Lists as filling uses them, from whatever a host or the browser's storage
 * gave: entries are text, trimmed, without blanks or repeats, and a list
 * has at least two.
 */
export function cleanLists(lists: unknown): string[][] {
  if (!Array.isArray(lists)) return [];
  return distinct(
    lists.flatMap((list) => {
      if (!Array.isArray(list)) return [];
      const seen = new Set<string>();
      const entries = list.flatMap((e) => {
        const text = typeof e === 'string' || typeof e === 'number' ? String(e).trim() : '';
        if (!text || seen.has(text.toLowerCase())) return [];
        seen.add(text.toLowerCase());
        return [text];
      });
      return entries.length >= 2 ? [entries] : [];
    }),
  );
}

/** Entries as typed in the Custom Lists dialog: one to a line, or separated by commas, as Excel takes them. */
export function splitEntries(text: string): string[] {
  return text
    .split(/[\r\n,]+/)
    .map((e) => e.trim())
    .filter(Boolean);
}

const KEY = 'wbe-custom-lists';
/** Each user of the host has their own lists; without a user, the browser has one set. */
const keyFor = (user: string | undefined) => (user ? `${KEY}:${user}` : KEY);

/** The user's own lists, kept in this browser; none where its storage can't be read. */
export function loadUserLists(user: string | undefined): string[][] {
  try {
    const raw = localStorage.getItem(keyFor(user));
    return raw ? cleanLists(JSON.parse(raw)) : [];
  } catch {
    return [];
  }
}

/** Keeps the user's lists in this browser; false where its storage can't be written, as in some private windows. */
export function saveUserLists(user: string | undefined, lists: string[][]): boolean {
  try {
    if (lists.length) localStorage.setItem(keyFor(user), JSON.stringify(lists));
    else localStorage.removeItem(keyFor(user));
    return true;
  } catch {
    return false;
  }
}

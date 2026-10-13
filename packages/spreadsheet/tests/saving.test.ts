import { describe, expect, it } from 'vitest';
import { parseInput } from '../src/core/input';
import { localeInfo, normalizeForLocale, readLocalDate } from '../src/core/locale';
import { english } from '../src/i18n';
import { pasteRisks } from '../src/state/pasteRisks';
import { open, sheetIndex } from './helpers';
import { parseA1 } from '@ruled/formula/refs';

describe('locale-aware input', () => {
  it('knows which way round each locale writes dates and decimals', () => {
    expect(localeInfo('en-US')).toEqual({ dayFirst: false, decimalComma: false });
    expect(localeInfo('en-GB')).toEqual({ dayFirst: true, decimalComma: false });
    expect(localeInfo('de-DE')).toEqual({ dayFirst: true, decimalComma: true });
  });

  it('reads slash dates in the locale order and flags the ambiguous ones', () => {
    expect(readLocalDate('3/4/2026', localeInfo('en-GB'))).toMatchObject({ day: 3, month: 4, year: 2026, ambiguous: true });
    expect(readLocalDate('3/4/2026', localeInfo('en-US'))).toMatchObject({ day: 4, month: 3, ambiguous: true });
    expect(readLocalDate('25/12/26', localeInfo('en-GB'))).toMatchObject({ day: 25, month: 12, year: 2026, ambiguous: false });
    expect(readLocalDate('31/02/2026', localeInfo('en-GB'))).toBeUndefined();
  });

  it('stores typed dates and decimal-comma numbers correctly', () => {
    const gb = parseInput('3/4/2026', 'General', 'en-GB');
    expect(gb.content).toEqual({ kind: 'number', value: 46115 });
    expect(gb.numFmt).toBe('dd/mm/yyyy');
    expect(parseInput('3/4/2026', 'General', 'en-US').content).toEqual({ kind: 'number', value: 46085 });
    expect(parseInput('1.234,5', 'General', 'de-DE').content).toEqual({ kind: 'number', value: 1234.5 });
    expect(parseInput('12,5%', 'General', 'de-DE').content).toEqual({ kind: 'number', value: 0.125 });
    expect(normalizeForLocale('1,234.5', localeInfo('en-US')).text).toBe('1,234.5');
  });
});

describe('strict paste', () => {
  it('finds values that would not read back as pasted', () => {
    const risks = pasteRisks([['03/04/2026', '25/12/2026', '007', '12345678901234567890', 'plain', '42']], 'en-GB');
    expect(risks.map((r) => [r.text, english.text(r.reason), r.becomes])).toEqual([
      ['03/04/2026', 'ambiguous date', '3 Apr 2026'],
      ['007', 'leading zeros', '7'],
      ['12345678901234567890', 'too many digits', '12345678901234500000'],
    ]);
  });
});

describe('net changes', () => {
  it('lists each changed cell with its value before and after', () => {
    const s = open('kitchen-sink.xlsx');
    const data = sheetIndex(s, 'Data');
    s.apply({ type: 'input', sheet: data, ...parseA1('D2')!, text: '500' });
    s.apply({ type: 'input', sheet: data, ...parseA1('F3')!, text: '=D3*2' });
    s.apply({ type: 'input', sheet: data, ...parseA1('J9')!, text: 'note' });
    const before = String(s.model.sheets[data].scan!.grid.nums[s.model.sheets[data].scan!.grid.rowPtr[1] + 3]);
    expect(s.changes()).toEqual([
      { sheet: 'Data', ref: 'D2', before, after: '500' },
      { sheet: 'Data', ref: 'F3', before: '=D3*E3', after: '=D3*2' },
      { sheet: 'Data', ref: 'J9', before: '', after: 'note' },
    ]);
  });
});

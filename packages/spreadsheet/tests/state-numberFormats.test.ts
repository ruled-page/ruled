import { describe, expect, it } from 'vitest';
import { numberFormatCode, numberFormatKind, stepDecimals } from '../src/state/numberFormats';

describe('number formats', () => {
  it('give each kind its code for the locale, and name a code by its kind', () => {
    expect(numberFormatCode('currency', 'en-US')).toBe('"$"#,##0.00');
    expect(numberFormatCode('currency', 'en-GB')).toBe('"£"#,##0.00');
    expect(numberFormatCode('currency', 'de-DE')).toBe('#,##0.00 "€"');
    expect(numberFormatCode('currency', 'pt-BR')).toBe('"$"#,##0.00');
    expect(numberFormatCode('currency', 'not a locale')).toBe('"$"#,##0.00');
    expect(numberFormatCode('shortDate', 'en-US')).toBe('m/d/yyyy');
    expect(numberFormatCode('shortDate', 'en-GB')).toBe('dd/mm/yyyy');
    expect(numberFormatKind('0.00%', 'en-US')).toBe('percent');
    expect(numberFormatKind('0.0%', 'en-US')).toBeUndefined();
  });

  it('add and remove decimal places in every section, leaving text, colours and dates alone', () => {
    expect(stepDecimals('General', 1)).toBe('0.0');
    expect(stepDecimals('General', -1)).toBe('0');
    expect(stepDecimals('0', 1)).toBe('0.0');
    expect(stepDecimals('0', -1)).toBe('0');
    expect(stepDecimals('#,##0.00', 1)).toBe('#,##0.000');
    expect(stepDecimals('#,##0.0', -1)).toBe('#,##0');
    expect(stepDecimals('0.00%', -1)).toBe('0.0%');
    expect(stepDecimals('"$"#,##0.00;[Red]-"$"#,##0.00', 1)).toBe('"$"#,##0.000;[Red]-"$"#,##0.000');
    expect(stepDecimals('#,##0.00 "kr. 0"', -1)).toBe('#,##0.0 "kr. 0"');
    expect(stepDecimals('0.00E+00', -1)).toBe('0.0E+00');
    expect(stepDecimals('m/d/yyyy', 1)).toBe('m/d/yyyy');
    expect(stepDecimals('@', 1)).toBe('@');
    expect(stepDecimals('\\0 "x"', 1)).toBe('\\0 "x"');
  });
});

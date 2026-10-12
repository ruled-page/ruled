import { describe, expect, it } from 'vitest';
import { listedFormatCodes, numberFormatCode, numberFormatKind, stepDecimals, usableFormat } from '../src/state/numberFormats';

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

  it("list Excel's built-in codes for the locale, then the workbook's own, each once", () => {
    const us = listedFormatCodes(['General', '0.00', '0.0 "kg"', '', '0.0 "kg"'], 'en-US');
    expect(us.slice(0, 5)).toEqual(['General', '0', '0.00', '#,##0', '#,##0.00']);
    expect(us).toContain('m/d/yyyy');
    expect(us.at(-1)).toBe('0.0 "kg"');
    expect(us.filter((c) => c === '0.0 "kg"' || c === '0.00')).toHaveLength(2);
    expect(listedFormatCodes([], 'en-GB')).toContain('dd/mm/yyyy');
  });

  it('tell codes Excel could use from ones it could not', () => {
    expect(usableFormat('0.00;[Red]-0.00')).toBe(true);
    expect(usableFormat('[h]:mm')).toBe(true);
    expect(usableFormat('')).toBe(false);
    expect(usableFormat('"abc')).toBe(false);
    expect(usableFormat('0;0;0;0;0')).toBe(false);
  });
});

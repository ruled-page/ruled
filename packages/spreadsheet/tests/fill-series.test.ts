import { afterEach, describe, expect, it, vi } from 'vitest';
import { planFill, seriesLine, type FillSource } from '../src/core/fill';
import type { SeriesOptions } from '../src/core/types';
import { builtInLists, cleanLists, loadUserLists, namesIn, saveUserLists, splitEntries } from '../src/state/fillLists';

/** A day as Excel numbers it, in the 1900 or 1904 date system. */
const serial = (y: number, m: number, d: number, date1904 = false) => (Date.UTC(y, m - 1, d) - (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30))) / 86_400_000;

const series = (options: Partial<SeriesOptions> = {}): Omit<SeriesOptions, 'by'> => ({ type: 'linear', dateUnit: 'day', step: 1, trend: false, ...options });
/** The values written along a line of `length` cells that starts with `starts`. */
const line = (starts: number[], length: number, options: Partial<SeriesOptions> = {}, date1904 = false) => seriesLine(starts, length, series(options), date1904);

describe('Fill Series of numbers', () => {
  it('adds the step to the first number for the rest of the line', () => {
    expect(line([1], 5, { step: 2 })).toEqual({ from: 1, values: [3, 5, 7, 9], direction: 1 });
    expect(line([0], 4, { step: 0.1 })?.values).toEqual([0.1, 0.2, 0.3]);
    expect(line([5, 99], 3)?.values).toEqual([6, 7]);
  });

  it('multiplies by the step for a growth series', () => {
    expect(line([2], 4, { type: 'growth', step: 3 })?.values).toEqual([6, 18, 54]);
    expect(line([64], 4, { type: 'growth', step: 0.5 })?.values).toEqual([32, 16, 8]);
  });

  it('ends before a value that passes the stop value, leaving the cells after it', () => {
    expect(line([1], 10, { step: 1, stop: 4 })?.values).toEqual([2, 3, 4]);
    expect(line([2], 10, { type: 'growth', step: 3, stop: 20 })?.values).toEqual([6, 18]);
    expect(line([10], 10, { step: 1, stop: 5 })?.values).toEqual([]);
  });

  it('counts down with a negative step, ending below the stop value', () => {
    expect(line([10], 10, { step: -3, stop: 0 })).toEqual({ from: 1, values: [7, 4, 1], direction: -1 });
    expect(line([10], 4, { step: -3 })?.values).toEqual([7, 4, 1]);
  });

  it('repeats the first number with a step of nothing, whatever the stop', () => {
    expect(line([4], 3, { step: 0, stop: 10 })).toEqual({ from: 1, values: [4, 4], direction: 0 });
  });

  it('fits a least-squares line to the numbers a trend starts from, and replaces them too', () => {
    expect(line([1, 3, 4], 5, { trend: true, step: 100 })).toEqual({ from: 0, values: [1.16666666666667, 2.66666666666667, 4.16666666666667, 5.66666666666667, 7.16666666666667], direction: 1 });
    expect(line([2, 4], 4, { trend: true })?.values).toEqual([2, 4, 6, 8]);
    // One number has no slope to follow.
    expect(line([5], 3, { trend: true })?.values).toEqual([5, 5, 5]);
  });

  it('fits an exponential curve for a growth trend, through numbers above zero only', () => {
    expect(line([1, 2, 4], 5, { type: 'growth', trend: true })?.values).toEqual([1, 2, 4, 8, 16]);
    expect(line([3, 0, 9], 5, { type: 'growth', trend: true })).toBeUndefined();
    expect(line([-1, -2], 5, { type: 'growth', trend: true })).toBeUndefined();
  });

  it('keeps the numbers a trend is fitted to whatever the stop value, and stops after them', () => {
    expect(line([1, 2, 3], 8, { trend: true, stop: 4 })?.values).toEqual([1, 2, 3, 4]);
    expect(line([1, 2, 3], 8, { trend: true, stop: 1 })?.values).toEqual([1, 2, 3]);
  });
});

describe('Fill Series of dates', () => {
  const date = (options: Partial<SeriesOptions>) => ({ type: 'date' as const, ...options });

  it('steps days, keeping the time of day', () => {
    expect(line([serial(2026, 1, 1)], 3, date({ step: 7 }))?.values).toEqual([serial(2026, 1, 8), serial(2026, 1, 15)]);
    expect(line([serial(2026, 1, 1) + 0.5], 2, date({ dateUnit: 'month' }))?.values).toEqual([serial(2026, 2, 1) + 0.5]);
  });

  it('steps weekdays, skipping Saturdays and Sundays, from a weekend too', () => {
    const friday = serial(2026, 1, 2);
    expect(line([friday], 4, date({ dateUnit: 'weekday' }))?.values).toEqual([serial(2026, 1, 5), serial(2026, 1, 6), serial(2026, 1, 7)]);
    expect(line([serial(2026, 1, 3)], 3, date({ dateUnit: 'weekday' }))?.values).toEqual([serial(2026, 1, 5), serial(2026, 1, 6)]);
    expect(line([friday], 3, date({ dateUnit: 'weekday', step: 6 }))?.values).toEqual([serial(2026, 1, 12), serial(2026, 1, 20)]);
    expect(line([serial(2026, 1, 5)], 3, date({ dateUnit: 'weekday', step: -1 }))?.values).toEqual([friday, serial(2026, 1, 1)]);
  });

  it('steps months from the first date, ending shorter months on their last day', () => {
    expect(line([serial(2026, 1, 31)], 4, date({ dateUnit: 'month' }))?.values).toEqual([serial(2026, 2, 28), serial(2026, 3, 31), serial(2026, 4, 30)]);
    expect(line([serial(2026, 3, 15)], 3, date({ dateUnit: 'month', step: -2 }))?.values).toEqual([serial(2026, 1, 15), serial(2025, 11, 15)]);
  });

  it('steps years, from 29 February to the 28th and back in leap years', () => {
    expect(line([serial(2024, 2, 29)], 5, date({ dateUnit: 'year' }))?.values).toEqual([serial(2025, 2, 28), serial(2026, 2, 28), serial(2027, 2, 28), serial(2028, 2, 29)]);
  });

  it('stops before a date past the stop date', () => {
    expect(line([serial(2026, 1, 31)], 12, date({ dateUnit: 'month', stop: serial(2026, 4, 15) }))?.values).toEqual([serial(2026, 2, 28), serial(2026, 3, 31)]);
  });

  it('counts days in the 1904 date system as the workbook does', () => {
    const in1904 = (y: number, m: number, d: number) => serial(y, m, d, true);
    expect(line([in1904(2026, 1, 31)], 3, date({ dateUnit: 'month' }), true)?.values).toEqual([in1904(2026, 2, 28), in1904(2026, 3, 31)]);
    expect(line([in1904(2026, 1, 2)], 2, date({ dateUnit: 'weekday' }), true)?.values).toEqual([in1904(2026, 1, 5)]);
    expect(line([in1904(2024, 2, 29)], 2, date({ dateUnit: 'year' }), true)?.values).toEqual([in1904(2025, 2, 28)]);
  });

  it('ends at the last date Excel shows', () => {
    expect(line([serial(9999, 12, 30)], 5, date({}))?.values).toEqual([serial(9999, 12, 31)]);
  });
});

describe('lists that filling continues', () => {
  const s = (...texts: string[]): FillSource[] => texts.map((value) => ({ content: { kind: 'string', value }, isDate: false }));
  const filled = (sources: FillSource[], positions: number[], lists?: string[][]) =>
    planFill(sources, positions, false, lists).map((p) => (p.kind === 'copy' ? `copy ${p.source}` : (p.content as { value: unknown }).value));
  const compass = [['North', 'East', 'South', 'West']];

  it('continues a custom list from any entry, matching whatever its case and writing in that case', () => {
    expect(filled(s('South'), [1, 2], compass)).toEqual(['West', 'North']);
    expect(filled(s('south'), [1], compass)).toEqual(['west']);
    expect(filled(s('EAST'), [1], compass)).toEqual(['SOUTH']);
    expect(filled(s('North', 'South'), [2, 3], compass)).toEqual(['North', 'South']);
    expect(filled(s('South'), [-1], compass)).toEqual(['East']);
  });

  it('writes a list of small letters with a capital first when the cell starts with one', () => {
    expect(filled(s('Alpha'), [1], [['alpha', 'beta', 'gamma']])).toEqual(['Beta']);
    expect(filled(s('McKinley'), [1], [['mckinley', 'denali']])).toEqual(['denali']);
  });

  it('names days and months in the language, short and full, with and without the full stop', () => {
    const german = builtInLists(['de']);
    expect(namesIn('de')).toContainEqual(['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember']);
    expect(filled(s('Montag'), [1, 6], german)).toEqual(['Dienstag', 'Sonntag']);
    expect(filled(s('Mo.'), [1], german)).toEqual(['Di.']);
    expect(filled(s('Mo'), [1], german)).toEqual(['Di']);
    expect(filled(s('Jan.'), [1], german)).toEqual(['Feb.']);
    expect(filled(s('MÄRZ'), [1], german)).toEqual(['APR']);
    expect(filled(s('Juni'), [1], german)).toEqual(['Juli']);
    // And as German Excel writes them.
    expect(filled(s('Mrz'), [1, 10], german)).toEqual(['Apr', 'Jan']);
    expect(filled(s('Mo.', 'Mi.'), [2], german)).toEqual(['Fr.']);
  });

  it('keeps English working in every language, after the language’s own names', () => {
    const german = builtInLists(['de', 'de-DE']);
    expect(filled(s('Mon'), [1], german)).toEqual(['Tue']);
    expect(filled(s('November'), [2], german)).toEqual(['Januar']);
    expect(filled(s('November'), [2], builtInLists(['en']))).toEqual(['January']);
    expect(filled(s('Nov'), [1, 2])).toEqual(['Dec', 'Jan']);
    expect(builtInLists(['en', 'en-US']).length).toBe(4);
  });

  it('takes text entries, trimmed, without blanks or repeats, two or more to a list', () => {
    expect(cleanLists([[' a ', 'b', '', 'A', 3], ['only'], 'x', [['nested']], ['A', 'B', '3']])).toEqual([['a', 'b', '3']]);
    expect(cleanLists(undefined)).toEqual([]);
    expect(splitEntries('North\r\nEast, South\n\nWest,')).toEqual(['North', 'East', 'South', 'West']);
  });
});

describe("the user's own lists", () => {
  afterEach(() => vi.unstubAllGlobals());

  function memoryStorage() {
    const items = new Map<string, string>();
    return {
      items,
      getItem: (k: string) => items.get(k) ?? null,
      setItem: (k: string, v: string) => void items.set(k, v),
      removeItem: (k: string) => void items.delete(k),
    };
  }

  it('are kept in the browser for each user', () => {
    const storage = memoryStorage();
    vi.stubGlobal('localStorage', storage);
    expect(saveUserLists('ana', [['x', 'y']])).toBe(true);
    expect(saveUserLists(undefined, [['p', 'q']])).toBe(true);
    expect(loadUserLists('ana')).toEqual([['x', 'y']]);
    expect(loadUserLists('ben')).toEqual([]);
    expect(loadUserLists(undefined)).toEqual([['p', 'q']]);
    saveUserLists('ana', []);
    expect(loadUserLists('ana')).toEqual([]);
    storage.items.set('wbe-custom-lists', '{not json');
    expect(loadUserLists(undefined)).toEqual([]);
  });

  it('come back empty, and say they were not kept, where storage refuses', () => {
    vi.stubGlobal('localStorage', {
      getItem: () => {
        throw new Error('denied');
      },
      setItem: () => {
        throw new Error('denied');
      },
    });
    expect(loadUserLists('ana')).toEqual([]);
    expect(saveUserLists('ana', [['x', 'y']])).toBe(false);
  });
});

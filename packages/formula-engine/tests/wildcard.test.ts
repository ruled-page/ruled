import { describe, expect, it } from 'vitest';
import { hasWildcards, matchesWhole, matchSpan, parsePattern, searchIn } from '../src/wildcard';

describe("Excel's wildcards", () => {
  it('match whole texts and find spans inside them', () => {
    const p = parsePattern('a*c?');
    expect(matchesWhole(p, 'abbbcd')).toBe(true);
    expect(matchesWhole(p, 'abbbc')).toBe(false);
    expect(matchSpan(p, 'xxabcdabcd')).toEqual([2, 6]);
    expect(matchSpan(p, 'xxabcdabcd', 3)).toEqual([6, 10]);
    expect(matchSpan(parsePattern('b*'), 'abcb')).toEqual([1, 4]);
    expect(matchSpan(parsePattern('*'), 'abc', 1)).toEqual([1, 3]);
    expect(searchIn(parsePattern('~*x'), 'a*x')).toBe(1);
    expect(searchIn(parsePattern('q'), 'abc')).toBe(-1);
    expect([hasWildcards('a?'), hasWildcards('a~'), hasWildcards('abc')]).toEqual([true, true, false]);
  });

  it('take time in proportion to the text, whatever the pattern', () => {
    const text = 'a'.repeat(20_000);
    const t0 = performance.now();
    expect(matchesWhole(parsePattern('*a*a*a*a*a*b'), text)).toBe(false);
    expect(matchSpan(parsePattern('*a*a*a*a*a*b'), text)).toBeUndefined();
    expect(performance.now() - t0).toBeLessThan(2000);
  });
});

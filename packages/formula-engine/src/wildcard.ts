// Excel's wildcards, matched in time proportional to the text's length times
// the pattern's, as regular expressions built from them can backtrack for
// minutes on long text. The engine's criteria and the editor's find use them.

const LITERAL = 0;
const ANY = 1;
const STAR = 2;

export interface Pattern {
  kinds: number[];
  chars: string[];
}

/** Excel's wildcards: * for any run, ? for any character, ~ before either to mean it literally. */
export function parsePattern(p: string): Pattern {
  const kinds: number[] = [];
  const chars: string[] = [];
  for (let i = 0; i < p.length; i++) {
    let c = p[i];
    let kind = c === '*' ? STAR : c === '?' ? ANY : LITERAL;
    if (c === '~' && i < p.length - 1) {
      c = p[++i];
      kind = LITERAL;
    }
    // Runs of * mean the same as one.
    if (kind === STAR && kinds[kinds.length - 1] === STAR) continue;
    kinds.push(kind);
    chars.push(c);
  }
  return { kinds, chars };
}

/** Whether text has wildcards (or the ~ that escapes them) to read as a pattern. */
export const hasWildcards = (text: string) => /[*?~]/.test(text);

/** Whether the whole of s matches, in at most length × pattern-length steps. */
export function matchesWhole({ kinds, chars }: Pattern, s: string): boolean {
  let si = 0;
  let pi = 0;
  let star = -1;
  let mark = 0;
  while (si < s.length) {
    if (pi < kinds.length && (kinds[pi] === ANY || (kinds[pi] === LITERAL && chars[pi] === s[si]))) {
      si++;
      pi++;
    } else if (pi < kinds.length && kinds[pi] === STAR) {
      star = pi++;
      mark = si;
    } else if (star !== -1) {
      pi = star + 1;
      si = ++mark;
    } else {
      return false;
    }
  }
  while (pi < kinds.length && kinds[pi] === STAR) pi++;
  return pi === kinds.length;
}

/** Earliest index at or after `from` where the pattern's tokens [a, b) match, or -1. */
function findSegment({ kinds, chars }: Pattern, a: number, b: number, s: string, from: number): number {
  for (let j = from; j + (b - a) <= s.length; j++) {
    let k = a;
    while (k < b && (kinds[k] === ANY || chars[k] === s[j + k - a])) k++;
    if (k === b) return j;
  }
  return -1;
}

/**
 * Where the pattern first matches inside s, at or after `from`: its start
 * and end. The pieces between stars are found one after another, each as
 * early as it can be, which is where the leftmost match puts them; a
 * trailing star takes the rest of the text.
 */
export function matchSpan(pattern: Pattern, s: string, from = 0): [number, number] | undefined {
  const { kinds } = pattern;
  let start = -1;
  let pos = from;
  let a = 0;
  while (a <= kinds.length) {
    let b = a;
    while (b < kinds.length && kinds[b] !== STAR) b++;
    if (b > a || start === -1) {
      const at = b > a ? findSegment(pattern, a, b, s, pos) : pos;
      if (at === -1) return undefined;
      if (start === -1) start = at;
      pos = at + (b - a);
    }
    a = b + 1;
  }
  return [start, kinds[kinds.length - 1] === STAR ? s.length : pos];
}

/** Where the pattern first matches inside s, or -1. */
export function searchIn(pattern: Pattern, s: string): number {
  return matchSpan(pattern, s)?.[0] ?? -1;
}

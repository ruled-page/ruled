import { parseA1Ref, type Token } from '@borgar/fx';
import { CellError, ErrorType, HyperFormula } from 'hyperformula';
import { MAX_COLS, MAX_ROWS } from './refs';

/**
 * Limits on what the calculation engine is asked to do, so that a workbook
 * can't hang the editor or run it out of memory with a few short formulas.
 * Formulas over the limits keep Excel's saved value, like any other formula
 * the engine can't calculate.
 */

/** Most cells one reference may cover: ten whole columns. */
export const MAX_RANGE_CELLS = 10 * MAX_ROWS;
/** Most values a function such as SEQUENCE may produce. */
export const MAX_ARRAY_CELLS = 1_000_000;

/** Functions whose literal arguments give the size of the array they return. */
const ARRAY_SIZE: Record<string, (args: number[]) => number> = {
  SEQUENCE: ([rows = 1, cols = 1]) => rows * cols,
  MUNIT: ([n = 1]) => n * n,
  RANDARRAY: ([rows = 1, cols = 1]) => rows * cols,
};

/** Whether a tokenized formula reads a range or builds an array too large to calculate here. */
export function tooLargeToCalculate(tokens: Token[]): boolean {
  const solid = tokens.filter((t) => t.type !== 'whitespace' && t.type !== 'newline');
  for (let i = 0; i < solid.length; i++) {
    const t = solid[i];
    if (t.type === 'range' || t.type === 'range_beam') {
      const r = parseA1Ref(t.value, { allowTernary: false });
      const range = r && 'range' in r ? r.range : undefined;
      if (!range) continue;
      const rows = (range.bottom ?? MAX_ROWS - 1) - (range.top ?? 0) + 1;
      const cols = (range.right ?? MAX_COLS - 1) - (range.left ?? 0) + 1;
      if (rows * cols > MAX_RANGE_CELLS) return true;
    }
    if (t.type === 'func') {
      const size = ARRAY_SIZE[t.value.toUpperCase().replace(/^_XLFN\./, '')];
      if (!size || solid[i + 1]?.value !== '(') continue;
      const args: number[] = [];
      for (let k = i + 2; k < solid.length && solid[k].type === 'number'; k += 2) {
        args.push(Math.trunc(Number(solid[k].value)));
        if (solid[k + 1]?.value !== ',') break;
      }
      if (size(args) > MAX_ARRAY_CELLS) return true;
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Wildcards

const LITERAL = 0;
const ANY = 1;
const STAR = 2;

interface Pattern {
  kinds: number[];
  chars: string[];
}

/** Excel's wildcards: * for any run, ? for any character, ~ before either to mean it literally. */
function parsePattern(p: string): Pattern {
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

/** Whether the whole of s matches, in at most length × pattern-length steps. */
function matchesWhole({ kinds, chars }: Pattern, s: string): boolean {
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
 * Where the pattern first matches inside s. The pieces between stars are
 * found one after another, each as early as it can be, which is where the
 * leftmost match puts them.
 */
function searchIn(pattern: Pattern, s: string): number {
  const { kinds } = pattern;
  let start = -1;
  let pos = 0;
  let a = 0;
  while (a <= kinds.length) {
    let b = a;
    while (b < kinds.length && kinds[b] !== STAR) b++;
    if (b > a || start === -1) {
      const at = b > a ? findSegment(pattern, a, b, s, pos) : pos;
      if (at === -1) return -1;
      if (start === -1) start = at;
      pos = at + (b - a);
    }
    a = b + 1;
  }
  return start;
}

let installed = false;

/**
 * Replaces the engine's wildcard matching, which turns "*a*a*b" into a
 * regular expression that backtracks for minutes on a long text, with
 * matching that takes at most text length × pattern length steps, and caps
 * SEQUENCE at MAX_ARRAY_CELLS values. Both reach into the engine's internals,
 * so tests check they still apply after an upgrade.
 */
export function installEngineLimits(): void {
  if (installed) return;
  installed = true;
  const probe = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  try {
    type Helper = {
      config: { useWildcards: boolean; useRegularExpressions: boolean; matchWholeCell: boolean };
      normalizeString(s: string): string;
      buildRegex(pattern: string, matchWholeCell?: boolean): unknown;
    };
    const helper = (probe as unknown as { _evaluator: { interpreter: { arithmeticHelper: Helper } } })._evaluator.interpreter.arithmeticHelper;
    const proto = Object.getPrototypeOf(helper) as Helper;
    const original = proto.buildRegex;
    proto.buildRegex = function (this: Helper, raw: string, matchWholeCell = true) {
      if (!this.config.useWildcards || this.config.useRegularExpressions) return original.call(this, raw, matchWholeCell);
      const pattern = parsePattern(this.normalizeString(raw));
      const whole = this.config.matchWholeCell && matchWholeCell;
      return {
        test: (s: string) => (whole ? matchesWhole(pattern, s) : searchIn(pattern, s) !== -1),
        exec: (s: string) => {
          if (whole) return matchesWhole(pattern, s) ? { index: 0 } : null;
          const index = searchIn(pattern, s);
          return index === -1 ? null : { index };
        },
      };
    };
  } finally {
    probe.destroy();
  }

  type Plugin = { prototype: Record<string, unknown> };
  const sequence = HyperFormula.getFunctionPlugin('SEQUENCE') as unknown as Plugin | undefined;
  const run = sequence?.prototype.sequence as ((this: unknown, ast: unknown, state: unknown) => unknown) | undefined;
  if (sequence && run) {
    sequence.prototype.sequence = function (this: { runFunction: (...a: unknown[]) => unknown; metadata: (n: string) => unknown }, ast: { args: unknown[] }, state: unknown) {
      // The arguments are worked out once here for the size, and again by the engine's own SEQUENCE.
      const tooBig = this.runFunction(ast.args, state, this.metadata('SEQUENCE'), (rows: number, cols: number) =>
        Math.trunc(rows) * Math.trunc(cols) > MAX_ARRAY_CELLS ? new CellError(ErrorType.NUM) : 0,
      );
      return tooBig instanceof CellError ? tooBig : run.call(this, ast, state);
    };
  }
}

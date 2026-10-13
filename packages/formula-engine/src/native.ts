import { parse, tokenize } from '@borgar/fx';
import type { EngineKind, Unsupported } from './api';
import { expandLets, stripFunctionPrefixes } from './formula';
import { tooLargeToCalculate } from './sizeLimits';
import { FUNCTIONS } from './functions';
import { NativeEngine } from './workbook';

/**
 * CELL's kinds of information this engine gives: those of the cell itself.
 * The rest need its number format, column width or protection, which only
 * the file has, and an address on another sheet needs the file's name.
 */
const CELL_INFO = new Set(['address', 'col', 'row', 'contents', 'type']);

type FxToken = ReturnType<typeof tokenize>[number];

/** Whether the CELL call at tokens[i] gives one of CELL_INFO, as a written text, for a reference it is given. */
function cellCalculable(tokens: FxToken[], i: number): boolean {
  const next = (j: number) => {
    while (++j < tokens.length && tokens[j].type === 'whitespace');
    return j;
  };
  const open = next(i);
  const info = next(open);
  const comma = next(info);
  const ref = next(comma);
  if (tokens[open]?.value !== '(' || tokens[info]?.type !== 'string' || tokens[comma]?.value !== ',') return false;
  const kind = tokens[info].value.slice(1, -1).toLowerCase();
  if (!CELL_INFO.has(kind)) return false;
  return kind !== 'address' || !tokens[ref]?.value.includes('!');
}

/**
 * The first function a parsed formula calls that the engine doesn't have,
 * leaving out LAMBDA parameters called as functions, and names defined as
 * LAMBDAs (`lambdas`, in capitals).
 */
function unknownCall(n: unknown, lambdas: ReadonlySet<string>, scope: ReadonlySet<string> = new Set()): string | undefined {
  if (!n || typeof n !== 'object') return undefined;
  const node = n as { type: string; callee?: { type: string; name?: string }; params?: { name: string }[]; body?: unknown; arguments?: unknown[]; elements?: unknown[][] };
  if (node.type === 'LambdaExpression') return unknownCall(node.body, lambdas, new Set([...scope, ...(node.params ?? []).map((p) => p.name.toUpperCase())]));
  if (node.type === 'CallExpression' && node.callee?.type === 'Identifier') {
    const name = (node.callee.name ?? '').toUpperCase();
    if ((!FUNCTIONS[name] || name === 'WBE.SPILL') && !scope.has(name) && !lambdas.has(name)) return name.replace(/^_XLUDF\./, '');
  } else if (node.type === 'CallExpression') {
    const inner = unknownCall(node.callee, lambdas, scope);
    if (inner) return inner;
  }
  for (const a of [...(node.arguments ?? []), ...(node.elements ?? []).flat()]) {
    const inner = unknownCall(a, lambdas, scope);
    if (inner) return inner;
  }
  return undefined;
}

/**
 * Why this engine can't calculate a formula (file form), or undefined when
 * it can: CELL for what it doesn't know, functions it doesn't have yet, links to other workbooks, and
 * formulas over the size limits.
 */
function unsupportedReason(fileFormula: string, lambdas: ReadonlySet<string> = new Set()): Unsupported | undefined {
  const expanded = expandLets(fileFormula);
  if (expanded === undefined) return { kind: 'function', name: 'LET' };
  let tokens;
  try {
    tokens = tokenize(stripFunctionPrefixes(expanded), { mergeRefs: true, allowTernary: false });
  } catch {
    return { kind: 'syntax' };
  }
  if (tooLargeToCalculate(tokens)) return { kind: 'tooLarge' };
  for (const t of tokens) {
    if (t.type === 'structured') return { kind: 'tableRefs' };
    if (t.type === 'func' && t.value.toUpperCase() === 'CELL' && !cellCalculable(tokens, tokens.indexOf(t))) return { kind: 'function', name: 'CELL' };
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_ternary' || t.type === 'range_named' || t.type === 'unknown') {
      if (t.value.includes('[')) return { kind: 'externalLinks' };
    }
  }
  let ast;
  try {
    ast = parse(tokens, { permitArrayRanges: true, permitArrayCalls: true });
  } catch {
    return { kind: 'syntax' };
  }
  const unknown = unknownCall(ast, lambdas);
  return unknown === undefined ? undefined : { kind: 'function', name: unknown };
}

/** The editor's own engine (src/engine): no dependencies but the formula parser, and Excel's rules throughout. */
export const nativeEngine: EngineKind = {
  id: 'native',
  name: 'Native',
  unsupportedReason,
  create: (sheets, names, date1904, lookups, iteration) => new NativeEngine(sheets, names, date1904, lookups, iteration),
};

/** The functions the engine calculates, by name in capitals. */
export const CALCULATED_FUNCTIONS: ReadonlySet<string> = new Set(Object.keys(FUNCTIONS));

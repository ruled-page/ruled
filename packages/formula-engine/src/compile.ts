import { fifteenDigits } from './excelNumbers';
import { parse, parseA1Ref, parseR1C1Ref, tokenize, translateFormulaToR1C1 } from '@borgar/fx';
import { stripFunctionPrefixes } from './formula';
import { MAX_COLS, MAX_ROWS, toA1 } from './refs';
import { FUNCTIONS } from './functions';
import { apply, attempt, each, pairwise, scalarOf, valuesOf, type Ctx, type Host } from './operands';
import { Areas, cancel, compare, err, isError, Lambda, Matrix, NAME, NULL, NUM, num, power, REF, Ref, text, VALUE, type Operand, type Scalar } from './values';

export type { Ctx, Host };

/**
 * Formulas compiled into JavaScript: each distinct formula, in its R1C1 form
 * relative to its own cell, is parsed once into a tree of closures, so a
 * column of copied formulas (=F2*G2, =F3*G3, …) shares one.
 */

type Node = (ctx: Ctx) => Operand;

/** A reference as written, relative to the formula's cell where it isn't absolute. */
export interface RefSpec {
  sheet: string | null;
  /** The last sheet of a three-dimensional reference, Jan:Mar!B2, whose first is `sheet`. */
  lastSheet?: string;
  r0: number | null;
  r1: number | null;
  c0: number | null;
  c1: number | null;
  $r0: boolean;
  $r1: boolean;
  $c0: boolean;
  $c1: boolean;
}

export interface Template {
  run: Node;
  /** The references it makes, for its dependencies. */
  refs: RefSpec[];
  /** Defined names it uses, in capitals, each qualified with a sheet as written ("Prices!RATE") when it is. */
  names: string[];
  /** Functions it calls, in capitals. */
  calls: Set<string>;
  volatile: boolean;
  /** Its result spills, wrapped by the session in WBE.SPILL(rows, cols, formula[, fixed]). */
  spill?: { rows: number; cols: number; fixed: boolean };
  /** Its outermost call is SUBTOTAL or AGGREGATE, which other subtotals leave out. */
  subtotal: boolean;
}

const cache = new Map<string, Template>();

/** The formula (with "=") in the cell at row, col, compiled; shared with every formula of the same relative form. */
export function compileAt(formula: string, row: number, col: number): Template {
  let r1c1: string;
  try {
    r1c1 = translateFormulaToR1C1(formula, toA1(row, col), {});
  } catch {
    return broken();
  }
  let t = cache.get(r1c1);
  if (!t) {
    t = compileR1C1(r1c1);
    if (cache.size > 50_000) cache.clear();
    cache.set(r1c1, t);
  }
  return t;
}

function broken(): Template {
  return { run: () => VALUE, refs: [], names: [], calls: new Set(), volatile: false, subtotal: false };
}

function compileR1C1(r1c1: string): Template {
  let ast;
  try {
    ast = parse(tokenize(r1c1, { r1c1: true, mergeRefs: true, allowTernary: false }), { permitArrayRanges: true, permitArrayCalls: true });
  } catch {
    return broken();
  }
  const t: Template = { run: () => null, refs: [], names: [], calls: new Set(), volatile: false, subtotal: false };
  t.run = node(ast as AstNode, t);
  const top = ast as AstNode;
  // A cell whose formula has a subtotal anywhere in it is one other subtotals leave out.
  if (t.calls.has('SUBTOTAL') || t.calls.has('AGGREGATE')) t.subtotal = true;
  if (top.type === 'CallExpression') {
    const name = callName(top);
    if (name === 'WBE.SPILL') {
      const lit = (a: AstNode | null | undefined) => (a && a.type === 'Literal' && typeof a.value === 'number' ? a.value : undefined);
      const rows = lit(top.arguments[0]);
      const cols = lit(top.arguments[1]);
      if (rows !== undefined && cols !== undefined) t.spill = { rows, cols, fixed: lit(top.arguments[3]) === 1 };
    }
  }
  return t;
}

// -----------------------------------------------------------------------------
// The tree

type AstNode =
  | { type: 'ReferenceIdentifier'; value: string; kind: string }
  | { type: 'Literal'; value: string | number | boolean; raw?: string }
  | { type: 'ErrorLiteral'; value: string }
  | { type: 'UnaryExpression'; operator: string; arguments: AstNode[] }
  | { type: 'BinaryExpression'; operator: string; arguments: AstNode[] }
  | { type: 'CallExpression'; callee: AstNode | { type: 'Identifier'; name: string }; arguments: (AstNode | null)[] }
  | { type: 'ArrayExpression'; elements: AstNode[][] }
  | { type: 'LambdaExpression'; params: { name: string }[]; body: AstNode | null }
  | { type: 'LetExpression' };

/** The function a call names; a LAMBDA called where it is written, LAMBDA(x,x+1)(2), has none. */
const callName = (n: { callee: object }) => ('name' in n.callee && typeof n.callee.name === 'string' ? stripFunctionPrefixes(n.callee.name).toUpperCase() : '');

/** LAMBDA parameters in scope while a body is compiled, innermost last. */
const scopes: Set<string>[] = [];
const inScope = (name: string) => scopes.some((s) => s.has(name));

/** How deeply LAMBDAs may call themselves (through names) before giving #NUM!, as Excel stops them. */
const MAX_LAMBDA_DEPTH = 1000;
let lambdaDepth = 0;

/** LAMBDA(x, y, body): a function, run with as many values as it has parameters, which see them by name. */
function lambda(n: { params: { name: string }[]; body: AstNode | null }, t: Template): Node {
  const params = n.params.map((p) => stripFunctionPrefixes(p.name).toUpperCase());
  scopes.push(new Set(params));
  let body: Node;
  try {
    body = node(n.body, t);
  } finally {
    scopes.pop();
  }
  return (ctx) =>
    new Lambda(params, (args) => {
      if (args.length !== params.length) return VALUE;
      if (lambdaDepth >= MAX_LAMBDA_DEPTH) return NUM;
      const vars = { ...ctx.vars };
      params.forEach((p, i) => (vars[p] = args[i]));
      lambdaDepth++;
      try {
        return body({ ...ctx, vars });
      } catch (e) {
        if (e instanceof RangeError) return NUM;
        throw e;
      } finally {
        lambdaDepth--;
      }
    });
}

/** Runs a LAMBDA with the arguments a call gives it; a name that isn't one is an error. */
function invoke(f: Operand | undefined, args: Node[], ctx: Ctx): Operand {
  if (f instanceof Lambda) return f.call(args.map((a) => a(ctx)));
  return isError(f) ? f : f === undefined ? NAME : VALUE;
}

function node(n: AstNode | null | undefined, t: Template): Node {
  if (!n) return () => null;
  switch (n.type) {
    case 'Literal': {
      // A number written with more than fifteen digits keeps fifteen, as Excel keeps them.
      const v = typeof n.value === 'number' && n.raw !== undefined ? fifteenDigits(n.raw) : n.value;
      return () => v;
    }
    case 'ErrorLiteral': {
      const e = err(n.value.toUpperCase());
      return () => e;
    }
    case 'ReferenceIdentifier':
      return reference(n, t);
    case 'ArrayExpression': {
      const rows = n.elements.map((row) => row.map((el) => node(el, t)));
      return (ctx) => Matrix.of(rows.map((row) => row.map((f) => scalarOf(f(ctx), ctx))));
    }
    case 'UnaryExpression':
      return unary(n.operator, node(n.arguments[0], t));
    case 'BinaryExpression':
      return binary(n.operator, node(n.arguments[0], t), node(n.arguments[1], t));
    case 'CallExpression':
      return call(n, t);
    case 'LambdaExpression':
      return lambda(n, t);
    default:
      // LET is written out before formulas reach the engine.
      return () => NAME;
  }
}

function reference(n: { value: string; kind: string }, t: Template): Node {
  if (n.kind === 'name') {
    if (n.value.includes('!')) {
      // A name qualified with a sheet, Prices!Rate: the one local to that sheet, else the workbook's.
      const ref = parseA1Ref(n.value, { allowTernary: false });
      const named = ref && 'name' in ref && ref.name && ref.context?.length === 1 ? { sheet: ref.context[0], name: ref.name } : undefined;
      if (!named || named.sheet.includes('[')) return () => REF;
      const { sheet } = named;
      const name = stripFunctionPrefixes(named.name).toUpperCase();
      t.names.push(`${sheet}!${name}`);
      return (ctx) => ctx.host.name(name, ctx, sheet);
    }
    const name = stripFunctionPrefixes(n.value).toUpperCase();
    // A parameter can hold a blank, REDUCE(,…) giving one.
    if (inScope(name)) return (ctx) => (ctx.vars && name in ctx.vars ? ctx.vars[name] : NAME);
    t.names.push(name);
    return (ctx) => ctx.host.name(name, ctx);
  }
  const parsed = parseR1C1Ref(n.value);
  if (!parsed || !('range' in parsed) || !parsed.range) return () => REF;
  const ctxNames = parsed.context ?? [];
  // A run of sheets, Jan:Mar!B2 (a sheet's name can't hold a colon).
  const [first, last] = ctxNames.length > 1 ? ctxNames : (ctxNames[0]?.split(':') ?? []);
  const r = parsed.range;
  const spec: RefSpec = {
    sheet: first ?? null,
    lastSheet: last,
    r0: r.r0 ?? null,
    r1: r.r1 ?? null,
    c0: r.c0 ?? null,
    c1: r.c1 ?? null,
    $r0: !!r.$r0,
    $r1: !!r.$r1,
    $c0: !!r.$c0,
    $c1: !!r.$c1,
  };
  t.refs.push(spec);
  return (ctx) => resolve(spec, ctx);
}

/**
 * A reference as written, made absolute at the formula's cell; #REF! off the
 * sheet or on a sheet not here. A three-dimensional one is the same cells on
 * each sheet from its first to its last, in the workbook's order.
 */
export function resolve(spec: RefSpec, ctx: { host: Host; sheet: number; row: number; col: number }): Ref | Areas | typeof REF {
  const sheet = spec.sheet === null ? ctx.sheet : ctx.host.sheetIndex(spec.sheet);
  if (sheet < 0) return REF;
  if (spec.lastSheet !== undefined) {
    const end = ctx.host.sheetIndex(spec.lastSheet);
    if (end < 0) return REF;
    const one = resolve({ ...spec, sheet: null, lastSheet: undefined }, { ...ctx, sheet: Math.min(sheet, end) });
    if (!(one instanceof Ref)) return one;
    const refs = [];
    for (let s = Math.min(sheet, end); s <= Math.max(sheet, end); s++) refs.push(new Ref(s, one.r1, one.c1, one.r2, one.c2));
    return new Areas(refs);
  }
  const at = (v: number, abs: boolean, here: number) => (abs ? v : here + v);
  let r1: number;
  let r2: number;
  if (spec.r0 === null && spec.r1 === null) {
    r1 = 0;
    r2 = MAX_ROWS - 1;
  } else {
    r1 = at(spec.r0 ?? spec.r1!, spec.r0 === null ? spec.$r1 : spec.$r0, ctx.row);
    r2 = at(spec.r1 ?? spec.r0!, spec.r1 === null ? spec.$r0 : spec.$r1, ctx.row);
  }
  let c1: number;
  let c2: number;
  if (spec.c0 === null && spec.c1 === null) {
    c1 = 0;
    c2 = MAX_COLS - 1;
  } else {
    c1 = at(spec.c0 ?? spec.c1!, spec.c0 === null ? spec.$c1 : spec.$c0, ctx.col);
    c2 = at(spec.c1 ?? spec.c0!, spec.c1 === null ? spec.$c0 : spec.$c1, ctx.col);
  }
  if (r1 > r2) [r1, r2] = [r2, r1];
  if (c1 > c2) [c1, c2] = [c2, c1];
  if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return REF;
  return new Ref(sheet, r1, c1, r2, c2);
}

function guard(f: (x: Scalar, y: Scalar) => Scalar): (x: Scalar, y: Scalar) => Scalar {
  return (x, y) => {
    try {
      return f(x, y);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  };
}

const numeric = (n: number) => (Number.isFinite(n) ? n : err('#NUM!'));

const ARITHMETIC: Record<string, (x: number, y: number) => Scalar> = {
  '+': (x, y) => cancel(x + y, x, y),
  '-': (x, y) => cancel(x - y, x, y),
  '*': (x, y) => x * y,
  '/': (x, y) => (y === 0 ? err('#DIV/0!') : x / y),
  '^': power,
};

const COMPARISON: Record<string, (d: number) => boolean> = {
  '=': (d) => d === 0,
  '<>': (d) => d !== 0,
  '<': (d) => d < 0,
  '>': (d) => d > 0,
  '<=': (d) => d <= 0,
  '>=': (d) => d >= 0,
};

function binary(op: string, left: Node, right: Node): Node {
  if (op === ':' || op === ' ' || op === ',') return referenceOperator(op, left, right);
  let f: (x: Scalar, y: Scalar) => Scalar;
  if (ARITHMETIC[op]) {
    const g = ARITHMETIC[op];
    f = guard((x, y) => {
      if (isError(x)) return x;
      if (isError(y)) return y;
      const r = g(num(x), num(y));
      return typeof r === 'number' ? numeric(r) : r;
    });
  } else if (COMPARISON[op]) {
    const test = COMPARISON[op];
    f = (x, y) => (isError(x) ? x : isError(y) ? y : test(compare(x, y)));
  } else if (op === '&') {
    // Cut off at what a cell holds, as Excel cuts it.
    f = guard((x, y) => (text(x) + text(y)).slice(0, 32767));
  } else {
    return () => err('#NAME?');
  }
  return (ctx) => pairwise(valuesOf(left(ctx), ctx), valuesOf(right(ctx), ctx), f);
}

function referenceOperator(op: string, left: Node, right: Node): Node {
  return (ctx) => {
    const a = left(ctx);
    const b = right(ctx);
    if (isError(a)) return a;
    if (isError(b)) return b;
    if (op === ',') {
      // A union, (A1:B2,D4): its areas in order, all on one sheet.
      if (!(a instanceof Ref || a instanceof Areas) || !(b instanceof Ref || b instanceof Areas)) return VALUE;
      const refs = [...(a instanceof Areas ? a.refs : [a]), ...(b instanceof Areas ? b.refs : [b])];
      return refs.every((r) => r.sheet === refs[0].sheet) ? new Areas(refs) : VALUE;
    }
    if (!(a instanceof Ref) || !(b instanceof Ref) || a.sheet !== b.sheet) return VALUE;
    if (op === ':') return new Ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
    if (op === ' ') {
      const r1 = Math.max(a.r1, b.r1);
      const c1 = Math.max(a.c1, b.c1);
      const r2 = Math.min(a.r2, b.r2);
      const c2 = Math.min(a.c2, b.c2);
      return r1 > r2 || c1 > c2 ? NULL : new Ref(a.sheet, r1, c1, r2, c2);
    }
    return VALUE;
  };
}

function unary(op: string, operand: Node): Node {
  switch (op) {
    case '-':
      return (ctx) => each(valuesOf(operand(ctx), ctx), (x) => (isError(x) ? x : -num(x)));
    case '+':
      return (ctx) => valuesOf(operand(ctx), ctx);
    case '%':
      return (ctx) => each(valuesOf(operand(ctx), ctx), (x) => (isError(x) ? x : num(x) / 100));
    case '@':
      return (ctx) => scalarOf(operand(ctx), ctx);
    case '#':
      return (ctx) => {
        const r = operand(ctx);
        if (!(r instanceof Ref) || !r.single) return REF;
        return ctx.host.spillOf(r.sheet, r.r1, r.c1) ?? REF;
      };
    default:
      return () => NAME;
  }
}

/** A function call: arguments evaluated (or, for IF and its kind, handed over to evaluate as needed), then the function. */
function call(n: { callee: AstNode | { type: 'Identifier'; name: string }; arguments: (AstNode | null)[] }, t: Template): Node {
  const args = n.arguments.map((a) => node(a, t));
  const name = callName(n);
  // LAMBDA(x,x+1)(2), or a call whose result is called.
  if (!name) {
    const f = node(n.callee as AstNode, t);
    return (ctx) => invoke(f(ctx), args, ctx);
  }
  // A parameter holding a LAMBDA, or a name defined as one.
  if (inScope(name)) return (ctx) => invoke(ctx.vars?.[name], args, ctx);
  t.calls.add(name);
  const spec = FUNCTIONS[name];
  if (!spec) {
    t.names.push(name);
    return (ctx) => invoke(ctx.host.name(name, ctx), args, ctx);
  }
  if (spec.volatile) t.volatile = true;
  if (args.length < spec.min || args.length > spec.max) return () => VALUE;
  if (spec.lazy) {
    const lazy = spec.lazy;
    return (ctx) => attempt(() => lazy(args, ctx));
  }
  return (ctx) => apply(spec, args.map((a) => a(ctx)), ctx);
}

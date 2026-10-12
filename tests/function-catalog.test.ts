import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { FUTURE_FUNCTIONS } from '@sitepen/formula-engine/formula';
import { functionHelp, functionsStartingWith, knownFunctionNames } from '../src/core/functionCatalog';
import { CALCULATED_FUNCTIONS } from '@sitepen/formula-engine/native';

/** The engine's own: the one spilling formulas are given to it in, and what A2# and @A1:A3 become in files. */
const ENGINE_ONLY = new Set(['WBE.SPILL', 'ANCHORARRAY', 'SINGLE']);

/** Excel functions the engine has no entry for: LET, which the editor writes out for it (expandLets), and LAMBDA, which it reads as part of the formula. */
const WRITTEN_OUT = new Set(['LET', 'LAMBDA']);

/** In Excel's file format but not offered by Excel as you type. */
const NOT_OFFERED = new Set(['ECMA.CEILING', 'QUERYSTRING']);

let registered: Set<string>;
beforeAll(() => {
  registered = new Set(CALCULATED_FUNCTIONS);
});

describe('the functions offered while typing a formula', () => {
  it("are marked calculated exactly when the worker's engine registers them, or the editor writes them out for it", () => {
    const names = knownFunctionNames();
    const calculated = names.filter((n) => functionHelp(n)!.calculated);
    expect(calculated.filter((n) => !registered.has(n) && !WRITTEN_OUT.has(n))).toEqual([]);
    expect(names.filter((n) => !functionHelp(n)!.calculated && registered.has(n))).toEqual([]);
    // Everything the engine calculates is offered, except its own functions.
    expect([...registered].filter((n) => !ENGINE_ONLY.has(n) && !calculated.includes(n))).toEqual([]);
    expect([...ENGINE_ONLY].filter((n) => !registered.has(n) || names.includes(n))).toEqual([]);
  });

  it("include every function Excel saves with the _xlfn. prefix, whether calculated or not", () => {
    expect([...FUTURE_FUNCTIONS].filter((n) => !NOT_OFFERED.has(n) && !functionHelp(n))).toEqual([]);
    expect(functionHelp('LAMBDA')).toMatchObject({ calculated: true });
    expect(functionHelp('CUBEVALUE')).toMatchObject({ calculated: false });
    expect(functionHelp('LET')).toMatchObject({ calculated: true });
    expect(functionHelp('xlookup')).toMatchObject({ name: 'XLOOKUP', calculated: true });
  });

  it('are in alphabetical order, each once', () => {
    const names = knownFunctionNames();
    expect(names.length).toBeGreaterThan(500);
    expect(names).toEqual([...new Set(names)].sort());
    expect(functionsStartingWith('vl').map((f) => f.name)).toEqual(['VLOOKUP']);
    expect(functionsStartingWith('SUMI').map((f) => f.name)).toEqual(['SUMIF', 'SUMIFS']);
    expect(functionsStartingWith('')).toEqual([]);
  });

  it('have arguments and a description for the common ones, each written well', () => {
    const source = readFileSync('src/core/functionCatalog.ts', 'utf8');
    const lines = source.slice(source.indexOf('const HELP = `'), source.indexOf('`;', source.indexOf('const HELP'))).split('\n').slice(1, -1);
    expect(lines.length).toBeGreaterThan(250);
    for (const line of lines) {
      const name = line.slice(0, line.indexOf('('));
      const fn = functionHelp(name);
      expect(fn?.name, line).toBe(name);
      expect(`${fn!.name}(${fn!.args}) ${fn!.description}`).toBe(line);
      // Arguments are names, optional ones in brackets, and "..." only last.
      expect(fn!.args, line).toMatch(/^((\[?[a-z_]+[a-z0-9_']*(, [a-z_]+[a-z0-9_']*)*\]?)(, (\[?[a-z_]+[a-z0-9_']*(, [a-z_]+[a-z0-9_']*)*\]?))*(, \.\.\.)?)?$/);
      expect(fn!.description, line).toMatch(/^[A-Z].*\.$/);
    }
    for (const common of ['SUM', 'IF', 'VLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'SUMIFS', 'COUNTIFS', 'TEXT', 'DATE', 'IFERROR', 'ROUND']) {
      expect(functionHelp(common)?.args, common).toBeDefined();
    }
    expect(functionHelp('BESSELJ')?.args).toBeUndefined();
  });
});

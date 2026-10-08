import { readFileSync } from 'node:fs';
import { HyperFormula } from 'hyperformula';
import { beforeAll, describe, expect, it } from 'vitest';
import { registerExcelFunctions } from '../src/core/excelFunctions';
import { FUTURE_FUNCTIONS } from '../src/core/formula';
import { functionHelp, functionsStartingWith, knownFunctionNames } from '../src/core/functionCatalog';

/** HyperFormula's own functions, which Excel doesn't have, so a formula using one would be #NAME? in Excel. */
const ENGINE_ONLY = new Set(
  (
    'ARRAYFORMULA ARRAY_CONSTRAIN CHIDISTRT CHIINVRT COUNTUNIQUE COVARIANCEP COVARIANCES FDISTRT FINVRT HF.ADD ' +
    'HF.CONCAT HF.DIVIDE HF.EQ HF.GT HF.GTE HF.LT HF.LTE HF.MINUS HF.MULTIPLY HF.NE HF.POW HF.UMINUS ' +
    'HF.UNARY_PERCENT HF.UPLUS INTERVAL ISBINARY LOGNORMINV MAXPOOL MEDIANPOOL POISSONDIST SKEWP SPLIT STDEVS ' +
    'TDIST2T TDISTRT TINV2T VARS VERSION WEIBULLDIST'
  ).split(' '),
);

/** Excel functions the engine lacks that the editor writes out for it (expandLets). */
const WRITTEN_OUT = new Set(['LET']);

/** In Excel's file format but not offered by Excel as you type. */
const NOT_OFFERED = new Set(['ECMA.CEILING', 'QUERYSTRING']);

let registered: Set<string>;
beforeAll(() => {
  // As the worker's engine has them.
  registerExcelFunctions();
  registered = new Set(HyperFormula.getRegisteredFunctionNames('enGB'));
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
    expect(functionHelp('LAMBDA')).toMatchObject({ calculated: false });
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

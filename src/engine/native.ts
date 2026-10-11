import { parse, tokenize } from '@borgar/fx';
import type { EngineKind, Unsupported } from '../core/engine';
import { expandLets, stripFunctionPrefixes } from '../core/formula';
import { tooLargeToCalculate } from '../core/sizeLimits';
import { FUNCTIONS } from './functions';
import { NativeEngine } from './workbook';

/**
 * Why this engine can't calculate a formula (file form), or undefined when
 * it can: functions it doesn't have yet, links to other workbooks, and
 * formulas over the size limits.
 */
function unsupportedReason(fileFormula: string): Unsupported | undefined {
  const expanded = expandLets(fileFormula);
  if (expanded === undefined) return { kind: 'function', name: 'LET' };
  let tokens;
  try {
    tokens = tokenize(expanded, { mergeRefs: true, allowTernary: false });
  } catch {
    return { kind: 'syntax' };
  }
  if (tooLargeToCalculate(tokens)) return { kind: 'tooLarge' };
  for (const t of tokens) {
    if (t.type === 'structured') return { kind: 'tableRefs' };
    if (t.type === 'func') {
      const name = stripFunctionPrefixes(t.value).toUpperCase();
      if (!FUNCTIONS[name] || name === 'WBE.SPILL') return { kind: 'function', name: name.replace(/^_XLUDF\./, '') };
    }
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_ternary' || t.type === 'unknown') {
      if (t.value.includes('[')) return { kind: 'externalLinks' };
    }
  }
  try {
    parse(tokens, { permitArrayRanges: true, permitArrayCalls: true });
  } catch {
    return { kind: 'syntax' };
  }
  return undefined;
}

/** The editor's own engine (src/engine): no dependencies but the formula parser, and Excel's rules throughout. */
export const nativeEngine: EngineKind = {
  id: 'native',
  name: 'Native',
  unsupportedReason,
  create: (sheets, names, date1904) => new NativeEngine(sheets, names, date1904),
};

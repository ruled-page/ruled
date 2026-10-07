import { parseA1Ref } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS, parseRange } from './refs';
import type { ListSource, ListValidation, RuleValidation, ValidationAlert } from './types';
import { attr, child, descendants, type XmlNode } from './xml';

export interface ListContext {
  /** Lower-cased sheet name → sheet index. */
  sheetIndex: Map<string, number>;
  ownSheet: number;
  names: { name: string; formula: string; localSheet?: number }[];
}

/** Values of an inline list, e.g. "North,South" or "Say ""hi"",Bye". */
function inlineValues(text: string): string[] {
  const body = text.slice(1, text.endsWith('"') ? -1 : undefined).replace(/""/g, '"');
  return body
    .split(',')
    .map((v) => v.trim())
    .filter((v) => v !== '');
}

export function resolveListSource(formula: string, ctx: ListContext): ListSource {
  let text = formula.trim().replace(/^=/, '');
  if (text.startsWith('"')) return { kind: 'values', values: inlineValues(text) };
  // A defined name: prefer one scoped to this sheet over a workbook-wide one.
  const lower = text.toLowerCase();
  const name =
    ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === ctx.ownSheet) ??
    ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === undefined);
  if (name) text = name.formula.trim().replace(/^=/, '');
  const ref = parseA1Ref(text, { allowTernary: false });
  if (!ref || !('range' in ref) || (ref.context?.length ?? 0) > 1) return { kind: 'formula', formula: formula.trim().replace(/^=/, '') };
  const sheetName = ref.context?.[0];
  const sheet = sheetName === undefined ? ctx.ownSheet : ctx.sheetIndex.get(sheetName.toLowerCase());
  if (sheet === undefined) return { kind: 'formula', formula: formula.trim().replace(/^=/, '') };
  const r = ref.range;
  return {
    kind: 'range',
    sheet,
    rect: {
      r1: r.top ?? 0,
      c1: r.left ?? 0,
      r2: r.bottom ?? (r.top === null ? MAX_ROWS - 1 : r.top),
      c2: r.right ?? (r.left === null ? MAX_COLS - 1 : r.left),
    },
  };
}

const KINDS = new Set(['whole', 'decimal', 'date', 'time', 'textLength', 'custom']);
const OPERATORS = new Set(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']);

/**
 * Data validation rules from a worksheet's non-cell XML. Excel writes them as
 * <dataValidation> or, for rules that point at other sheets in older files,
 * as <x14:dataValidation> inside extLst with the formulas and range as elements.
 */
export function parseValidations(meta: XmlNode, ctx: ListContext): { lists: ListValidation[]; rules: RuleValidation[]; other: number } {
  const lists: ListValidation[] = [];
  const rules: RuleValidation[] = [];
  let other = 0;
  for (const dv of descendants(meta, 'dataValidation')) {
    const type = attr(dv, 'type') ?? 'none';
    const formulaOf = (name: string) => {
      const node = child(dv, name);
      return (child(node, 'f')?.text ?? node?.text ?? '').trim();
    };
    const sqref = (attr(dv, 'sqref') ?? child(dv, 'sqref')?.text ?? '').trim();
    const rects = sqref
      .split(/\s+/)
      .map((r) => parseRange(r))
      .filter((r): r is NonNullable<typeof r> => !!r);
    const flag = (name: string, fallback: boolean) => {
      const v = attr(dv, name);
      return v === undefined ? fallback : v === '1' || v === 'true';
    };
    const style = attr(dv, 'errorStyle');
    const alert: ValidationAlert = {
      allowBlank: flag('allowBlank', false),
      showError: flag('showErrorMessage', false),
      errorStyle: style === 'warning' || style === 'information' ? style : 'stop',
      errorTitle: attr(dv, 'errorTitle') || undefined,
      error: attr(dv, 'error') || undefined,
      promptTitle: attr(dv, 'promptTitle') || undefined,
      prompt: attr(dv, 'prompt') || undefined,
    };
    const formula1 = formulaOf('formula1');
    if (!rects.length) continue;
    if (type === 'list' && formula1) {
      lists.push({
        ...alert,
        rects,
        source: resolveListSource(formula1, ctx),
        // Excel's showDropDown attribute is inverted: "1" hides the in-cell dropdown.
        inCellDropdown: !flag('showDropDown', false),
      });
    } else if (KINDS.has(type) && formula1) {
      const op = attr(dv, 'operator') ?? 'between';
      rules.push({
        ...alert,
        rects,
        kind: type as RuleValidation['kind'],
        operator: (OPERATORS.has(op) ? op : 'between') as RuleValidation['operator'],
        formula1: formula1.replace(/^=/, ''),
        formula2: formulaOf('formula2').replace(/^=/, '') || undefined,
      });
    } else {
      other++;
    }
  }
  return { lists, rules, other };
}


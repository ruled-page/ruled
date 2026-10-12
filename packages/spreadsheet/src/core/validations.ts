import { parseRange, rangeToA1 } from '@sitepen/formula-engine/refs';
import { resolveListSource, type ListContext } from './validationRules';
import { placeSheetElement } from './hyperlinks';
import type { ListSource, ListValidation, RuleValidation, ValidationAlert } from './types';

export { resolveListSource, subtractRect, withoutRects, type ListContext } from './validationRules';
import { attr, child, descendants, type XmlNode } from './xml';

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
      showPrompt: flag('showInputMessage', false),
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
        formula: formula1.replace(/^=/, ''),
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
    } else if (type === 'none') {
      // "Any value": nothing is checked, but an input message may show.
      rules.push({ ...alert, rects, kind: 'any', operator: 'between', formula1: '' });
    } else {
      other++;
    }
  }
  return { lists, rules, other };
}

const quoteAttr = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '&#10;');
const escapeText = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** One rule as Excel writes it in a worksheet's <dataValidations>. */
function validationXml(p: string, v: ListValidation | RuleValidation): string {
  const list = 'source' in v;
  const type = list ? 'list' : v.kind === 'any' ? 'none' : v.kind;
  const attrs: string[] = [`type="${type}"`];
  if (!list && v.kind !== 'any' && v.kind !== 'custom' && v.operator !== 'between') attrs.push(`operator="${v.operator}"`);
  if (v.errorStyle !== 'stop') attrs.push(`errorStyle="${v.errorStyle}"`);
  if (v.allowBlank) attrs.push('allowBlank="1"');
  // Excel's showDropDown is the other way round: "1" hides the in-cell dropdown.
  if (list && !v.inCellDropdown) attrs.push('showDropDown="1"');
  if (v.showPrompt ?? !!v.prompt) attrs.push('showInputMessage="1"');
  if (v.showError) attrs.push('showErrorMessage="1"');
  for (const [name, value] of [['errorTitle', v.errorTitle], ['error', v.error], ['promptTitle', v.promptTitle], ['prompt', v.prompt]] as const) {
    if (value) attrs.push(`${name}="${quoteAttr(value)}"`);
  }
  attrs.push(`sqref="${v.rects.map(rangeToA1).join(' ')}"`);
  const formula1 = list ? (v.formula ?? listFormula(v.source)) : v.formula1;
  const formula2 = list ? undefined : v.formula2;
  const formulas =
    (formula1 ? `<${p}formula1>${escapeText(formula1)}</${p}formula1>` : '') + (formula2 && (list || v.kind !== 'custom') ? `<${p}formula2>${escapeText(formula2)}</${p}formula2>` : '');
  return formulas ? `<${p}dataValidation ${attrs.join(' ')}>${formulas}</${p}dataValidation>` : `<${p}dataValidation ${attrs.join(' ')}/>`;
}

/** A list's source as a formula, where the file's own isn't known: its values quoted, or its range. */
function listFormula(source: ListSource): string {
  if (source.kind === 'values') return `"${source.values.map((v) => v.replace(/"/g, '""')).join(',')}"`;
  if (source.kind === 'formula') return source.formula;
  return rangeToA1(source.rect);
}

/**
 * The worksheet with its <dataValidations> as `lists` and `rules` have them:
 * in its place among the worksheet's elements, gone when there are none, and
 * the copies Excel 2010 keeps in the sheet's extensions taken out, as the
 * rules now in <dataValidations> stand in for them.
 */
export function writeValidations(xml: string, p: string, lists: readonly ListValidation[], rules: readonly RuleValidation[]): string {
  // Excel 2010's copies, in their own extension.
  xml = xml.replace(/<(\w+:)?ext\b[^>]*\buri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}"[^>]*>[\s\S]*?<\/\1?ext>/, '');
  xml = xml.replace(/<(\w+:)?extLst>\s*<\/\1?extLst>/, '');
  const all = [...lists, ...rules].filter((v) => v.rects.length);
  const el = all.length ? `<${p}dataValidations count="${all.length}">${all.map((v) => validationXml(p, v)).join('')}</${p}dataValidations>` : '';
  return placeSheetElement(xml, 'dataValidations', el);
}



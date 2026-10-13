import { displayFormula, shiftFormula } from '@ruled/formula/formula';
import { toA1 } from '@ruled/formula/refs';
import { parseInput } from '../core/input';
import type { CfRule, Cfvo, DiffStyle } from '../core/types';
import { msg, type MessageKey, type Text } from '../i18n/message';
import type { DraftPlace } from './validationDraft';

/** The rules Excel's Highlight Cells, Top/Bottom and New Rule menus make from a value or two, a formula, or nothing more. */
export type CfQuick =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equal'
  | 'textContains'
  | 'duplicate'
  | 'top'
  | 'topPercent'
  | 'bottom'
  | 'bottomPercent'
  | 'aboveAverage'
  | 'belowAverage'
  | 'formula';

/** How many values each asks for. */
export const CF_VALUES: Record<CfQuick, 0 | 1 | 2> = {
  greaterThan: 1,
  lessThan: 1,
  between: 2,
  equal: 1,
  textContains: 1,
  duplicate: 0,
  top: 1,
  topPercent: 1,
  bottom: 1,
  bottomPercent: 1,
  aboveAverage: 0,
  belowAverage: 0,
  formula: 1,
};

const red = { width: 1 as const, pattern: 'solid' as const, color: '#FF0000' };

/** The formats Excel's quick rules offer, in its order. */
export const CF_PRESETS: { key: MessageKey; style: DiffStyle }[] = [
  { key: 'cf.preset.lightRed', style: { fill: '#FFC7CE', color: '#9C0006' } },
  { key: 'cf.preset.yellow', style: { fill: '#FFEB9C', color: '#9C5700' } },
  { key: 'cf.preset.green', style: { fill: '#C6EFCE', color: '#006100' } },
  { key: 'cf.preset.lightRedFill', style: { fill: '#FFC7CE' } },
  { key: 'cf.preset.redText', style: { color: '#9C0006' } },
  { key: 'cf.preset.redBorder', style: { border: { left: red, right: red, top: red, bottom: red } } },
];

const anchorOf = (place: DraftPlace) => ({ row: Math.min(...place.rects.map((r) => r.r1)), col: Math.min(...place.rects.map((r) => r.c1)) });

/**
 * A value typed into a quick rule, as the rule keeps it: a number in the
 * locale, a formula moved from the active cell to the rule's top-left cell,
 * or text, quoted as Excel writes it.
 */
function valueFormula(text: string, place: DraftPlace): string | undefined {
  const t = text.trim();
  if (!t) return undefined;
  const parsed = parseInput(t, 'General', place.locale, place.date1904);
  if (parsed.error) return undefined;
  if (parsed.content.kind === 'number') return String(parsed.content.value);
  if (parsed.content.kind === 'formula') {
    const a = anchorOf(place);
    return shiftFormula(parsed.content.formula, place.active.row, place.active.col, a.row, a.col);
  }
  return `"${t.replace(/"/g, '""')}"`;
}

/**
 * The rule a quick dialog makes for `place`, with the format picked, or the
 * box whose value doesn't make one. Rules keep their formulas for the
 * top-left cell of their first range.
 */
export function quickRule(kind: CfQuick, values: string[], style: DiffStyle, place: DraftPlace): { rule: CfRule } | { problem: { message: MessageKey; box: number } } {
  const base = { rects: place.rects, priority: 0, stopIfTrue: false, formulas: [] as string[], style };
  const formulas = values.slice(0, CF_VALUES[kind]).map((v) => valueFormula(v, place));
  const missing = formulas.findIndex((f) => f === undefined);
  switch (kind) {
    case 'greaterThan':
    case 'lessThan':
    case 'equal':
    case 'between':
      if (missing >= 0) return { problem: { message: 'cf.problem.value', box: missing } };
      return { rule: { ...base, type: 'cellIs', operator: kind, formulas: formulas as string[] } };
    case 'textContains': {
      const text = values[0]?.trim() ?? '';
      if (!text) return { problem: { message: 'cf.problem.text', box: 0 } };
      const a = anchorOf(place);
      return { rule: { ...base, type: 'containsText', operator: 'containsText', text, formulas: [`NOT(ISERROR(SEARCH("${text.replace(/"/g, '""')}",${toA1(a.row, a.col)})))`] } };
    }
    case 'duplicate':
      return { rule: { ...base, type: 'duplicateValues' } };
    case 'top':
    case 'topPercent':
    case 'bottom':
    case 'bottomPercent': {
      const n = Number(values[0]?.trim());
      const percent = kind.endsWith('Percent');
      if (!Number.isInteger(n) || n < 1 || n > (percent ? 100 : 1000)) return { problem: { message: percent ? 'cf.problem.percent' : 'cf.problem.rank', box: 0 } };
      return { rule: { ...base, type: 'top10', rank: n, ...(percent && { percent: true }), ...(kind.startsWith('bottom') && { bottom: true }) } };
    }
    case 'aboveAverage':
    case 'belowAverage':
      return { rule: { ...base, type: 'aboveAverage', aboveAverage: kind === 'aboveAverage' } };
    case 'formula': {
      const t = values[0]?.trim() ?? '';
      const parsed = t ? parseInput(t.startsWith('=') ? t : `=${t}`, 'General', place.locale, place.date1904) : undefined;
      if (!parsed || parsed.error || parsed.content.kind !== 'formula') return { problem: { message: 'cf.problem.formula', box: 0 } };
      const a = anchorOf(place);
      const formula = shiftFormula(parsed.content.formula, place.active.row, place.active.col, a.row, a.col);
      if (!formula) return { problem: { message: 'cf.problem.formula', box: 0 } };
      return { rule: { ...base, type: 'expression', formulas: [formula] } };
    }
  }
}

/** The quick dialog a rule opens in to be changed, with its values as typed, or none for rules made another way. */
export function quickOf(rule: CfRule, place: DraftPlace): { kind: CfQuick; values: string[] } | undefined {
  const a = anchorOf(place);
  const shown = (f: string) => {
    if (/^-?\d+(\.\d+)?$/.test(f)) return Number(f).toLocaleString(place.locale, { useGrouping: false, maximumFractionDigits: 15 });
    if (/^"(?:[^"]|"")*"$/.test(f)) return f.slice(1, -1).replace(/""/g, '"');
    return displayFormula(shiftFormula(f, a.row, a.col, place.active.row, place.active.col) ?? f);
  };
  switch (rule.type) {
    case 'cellIs':
      return rule.operator === 'greaterThan' || rule.operator === 'lessThan' || rule.operator === 'equal' || rule.operator === 'between'
        ? { kind: rule.operator, values: rule.formulas.map(shown) }
        : undefined;
    case 'containsText':
      return rule.text !== undefined ? { kind: 'textContains', values: [rule.text] } : undefined;
    case 'duplicateValues':
      return { kind: 'duplicate', values: [] };
    case 'top10':
      return { kind: `${rule.bottom ? 'bottom' : 'top'}${rule.percent ? 'Percent' : ''}` as CfQuick, values: [String(rule.rank ?? 10)] };
    case 'aboveAverage':
      return rule.stdDev === undefined && !rule.equalAverage ? { kind: rule.aboveAverage === false ? 'belowAverage' : 'aboveAverage', values: [] } : undefined;
    case 'expression':
      return rule.formulas[0] ? { kind: 'formula', values: [shown(rule.formulas[0])] } : undefined;
    default:
      return undefined;
  }
}

const cfvo = (type: Cfvo['type'], val?: string): Cfvo => ({ type, ...(val !== undefined && { val }), gte: true });

/** Excel's Data Bars, Color Scales and Icon Sets menus: a rule for the selection, with Excel's usual thresholds. */
export const CF_GALLERY: {
  dataBars: { key: MessageKey; color: string }[];
  colorScales: { key: MessageKey; colors: string[] }[];
  iconSets: { key: MessageKey; set: string }[];
} = {
  dataBars: [
    { key: 'cf.bar.blue', color: '#638EC6' },
    { key: 'cf.bar.green', color: '#63C384' },
    { key: 'cf.bar.red', color: '#FF555A' },
    { key: 'cf.bar.orange', color: '#FFB628' },
    { key: 'cf.bar.lightBlue', color: '#008AEF' },
    { key: 'cf.bar.purple', color: '#D6007B' },
  ],
  colorScales: [
    { key: 'cf.scale.greenYellowRed', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
    { key: 'cf.scale.redYellowGreen', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
    { key: 'cf.scale.whiteGreen', colors: ['#FCFCFF', '#63BE7B'] },
    { key: 'cf.scale.whiteRed', colors: ['#FCFCFF', '#F8696B'] },
  ],
  iconSets: [
    { key: 'cf.icons.arrows', set: '3Arrows' },
    { key: 'cf.icons.trafficLights', set: '3TrafficLights1' },
    { key: 'cf.icons.flags', set: '3Flags' },
    { key: 'cf.icons.symbols', set: '3Symbols' },
  ],
};

/** A rule from Excel's Data Bars, Color Scales or Icon Sets galleries. */
export function galleryRule(rects: DraftPlace['rects'], pick: { color: string } | { colors: readonly string[] } | { set: string }): CfRule {
  const base = { rects, priority: 0, stopIfTrue: false, formulas: [] };
  if ('color' in pick) {
    return { ...base, type: 'dataBar', dataBar: { cfvo: [cfvo('min'), cfvo('max')], color: pick.color, showValue: true, gradient: true, minLength: 10, maxLength: 90 } };
  }
  if ('colors' in pick) {
    const points = pick.colors.length === 3 ? [cfvo('min'), cfvo('percentile', '50'), cfvo('max')] : [cfvo('min'), cfvo('max')];
    return { ...base, type: 'colorScale', colorScale: { cfvo: points, colors: [...pick.colors] } };
  }
  return { ...base, type: 'iconSet', iconSet: { set: pick.set, cfvo: [cfvo('percent', '0'), cfvo('percent', '33'), cfvo('percent', '67')], reverse: false, showValue: true } };
}

const OPERATOR_KEYS: Record<string, MessageKey> = {
  greaterThan: 'cf.describe.greaterThan',
  lessThan: 'cf.describe.lessThan',
  between: 'cf.describe.between',
  notBetween: 'cf.describe.notBetween',
  equal: 'cf.describe.equal',
  notEqual: 'cf.describe.notEqual',
  greaterThanOrEqual: 'cf.describe.greaterThanOrEqual',
  lessThanOrEqual: 'cf.describe.lessThanOrEqual',
};

/** A rule in a line, as Excel's Rules Manager lists it: "Cell value > 5", "Top 10", "Data bar". */
export function describeCfRule(rule: CfRule): Text {
  const [a = '', b = ''] = rule.formulas;
  switch (rule.type) {
    case 'cellIs':
      return OPERATOR_KEYS[rule.operator ?? ''] ? msg(OPERATOR_KEYS[rule.operator!], { a, b }) : msg('cf.describe.cellValue');
    case 'containsText':
      return msg('cf.describe.contains', { text: rule.text ?? '' });
    case 'duplicateValues':
      return msg('cf.describe.duplicate');
    case 'uniqueValues':
      return msg('cf.describe.unique');
    case 'top10':
      return msg(rule.bottom ? (rule.percent ? 'cf.describe.bottomPercent' : 'cf.describe.bottom') : rule.percent ? 'cf.describe.topPercent' : 'cf.describe.top', { count: rule.rank ?? 10 });
    case 'aboveAverage':
      return msg(rule.aboveAverage === false ? 'cf.describe.belowAverage' : 'cf.describe.aboveAverage');
    case 'expression':
      return msg('cf.describe.formula', { formula: displayFormula(a) });
    case 'colorScale':
      return msg('cf.describe.colorScale');
    case 'dataBar':
      return msg('cf.describe.dataBar');
    case 'iconSet':
      return msg('cf.describe.iconSet');
    default:
      return msg('cf.describe.other', { type: rule.type });
  }
}

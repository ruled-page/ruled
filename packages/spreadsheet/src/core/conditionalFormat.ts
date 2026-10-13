import type { EngineValue } from '@ruled/formula';
import { makeShifter } from '@ruled/formula/formula';
import { cellKey, MAX_COLS, parseRange, rangeToA1 } from '@ruled/formula/refs';
import { placeSheetElement } from './hyperlinks';
import type { StyleBook } from './styles';
import type { CfCell, CfResult, CfRule, Cfvo, DiffStyle, Rect } from './types';
import { attr, child, children, descendants, type XmlNode } from './xml';

// ---------------------------------------------------------------------------
// Reading rules

function readCfvo(node: XmlNode): Cfvo {
  const type = (attr(node, 'type') ?? 'num') as Cfvo['type'];
  // x14 thresholds keep their value in an <xm:f> child.
  const val = attr(node, 'val') ?? child(node, 'f')?.text;
  return { type, val: val?.trim(), gte: attr(node, 'gte') !== '0' };
}

const flag = (node: XmlNode | undefined, name: string, fallback: boolean) => {
  const v = attr(node, name);
  return v === undefined ? fallback : v === '1' || v === 'true';
};

/**
 * Conditional formatting rules from a worksheet's non-cell XML: the
 * <conditionalFormatting> blocks, plus Excel 2010's x14 blocks in extLst,
 * which either add options to a rule (data bar fill, negative colour) or
 * hold whole rules of their own (for example ones that refer to other sheets).
 */
export function parseConditionalFormats(meta: XmlNode, styles: StyleBook): CfRule[] {
  const rules: CfRule[] = [];
  const byExtId = new Map<string, CfRule>();
  const blocks = descendants(meta, 'conditionalFormatting');
  const ruleOf = (r: XmlNode, rects: Rect[], x14: boolean): CfRule => {
    const dxfId = attr(r, 'dxfId');
    const inline = child(r, 'dxf');
    const rule: CfRule = {
      rects,
      type: attr(r, 'type') ?? 'expression',
      priority: Number(attr(r, 'priority') ?? 0),
      stopIfTrue: flag(r, 'stopIfTrue', false),
      style: dxfId !== undefined ? styles.dxfs[Number(dxfId)] : inline ? styles.dxf(inline) : undefined,
      operator: attr(r, 'operator'),
      formulas: (x14 ? children(r, 'f') : children(r, 'formula')).map((f) => f.text.trim()),
      text: attr(r, 'text'),
    };
    if (dxfId !== undefined) rule.dxfId = Number(dxfId);
    if (attr(r, 'rank') !== undefined) rule.rank = Number(attr(r, 'rank'));
    if (flag(r, 'percent', false)) rule.percent = true;
    if (flag(r, 'bottom', false)) rule.bottom = true;
    if (rule.type === 'aboveAverage') {
      rule.aboveAverage = flag(r, 'aboveAverage', true);
      rule.equalAverage = flag(r, 'equalAverage', false);
      if (attr(r, 'stdDev') !== undefined) rule.stdDev = Number(attr(r, 'stdDev'));
    }
    const scale = child(r, 'colorScale');
    if (scale) {
      rule.colorScale = {
        cfvo: children(scale, 'cfvo').map(readCfvo),
        colors: children(scale, 'color').map((c) => styles.color(c) ?? '#FFFFFF'),
      };
    }
    const bar = child(r, 'dataBar');
    if (bar) {
      rule.dataBar = {
        cfvo: children(bar, 'cfvo').map(readCfvo),
        color: styles.color(child(bar, 'color') ?? child(bar, 'fillColor')) ?? '#638EC6',
        showValue: flag(bar, 'showValue', true),
        gradient: flag(bar, 'gradient', true),
        // Excel 2007 bars run from 10% to 90% of the cell; 2010 bars use the full width.
        minLength: Number(attr(bar, 'minLength') ?? (x14 ? 0 : 10)),
        maxLength: Number(attr(bar, 'maxLength') ?? (x14 ? 100 : 90)),
        negativeColor: styles.color(child(bar, 'negativeFillColor')),
      };
    }
    const icons = child(r, 'iconSet');
    if (icons) {
      rule.iconSet = {
        set: attr(icons, 'iconSet') ?? '3TrafficLights1',
        cfvo: children(icons, 'cfvo').map(readCfvo),
        reverse: flag(icons, 'reverse', false),
        showValue: flag(icons, 'showValue', true),
      };
    }
    return rule;
  };

  for (const block of blocks) {
    const x14 = block.qname.includes(':');
    const sqref = (attr(block, 'sqref') ?? child(block, 'sqref')?.text ?? '').trim();
    const rects = sqref
      .split(/\s+/)
      .map((s) => parseRange(s))
      .filter((r): r is Rect => !!r);
    if (!rects.length) continue;
    for (const r of children(block, 'cfRule')) {
      if (x14) {
        const base = byExtId.get(attr(r, 'id') ?? '');
        if (base) {
          // Options Excel 2010 added to a rule written in the main block.
          const extra = ruleOf(r, rects, true);
          if (base.dataBar && extra.dataBar) base.dataBar = { ...extra.dataBar, color: base.dataBar.color };
          if (base.iconSet && extra.iconSet) base.iconSet = extra.iconSet;
          continue;
        }
      }
      const rule = ruleOf(r, rects, x14);
      // A rule of Excel 2010's own, in the extensions only, goes by its id there.
      const extId = x14 ? attr(r, 'id') : descendants(child(r, 'extLst'), 'id')[0]?.text.trim();
      if (extId) {
        rule.extId = extId;
        byExtId.set(extId, rule);
      }
      if (x14) rule.x14Only = true;
      rules.push(rule);
    }
  }
  return rules.sort((a, b) => a.priority - b.priority);
}

// ---------------------------------------------------------------------------
// Evaluating rules

export interface CfContext {
  /** A cell's current value; null when empty. */
  value(row: number, col: number): EngineValue;
  /** Evaluates a formula in file form on this sheet. */
  calculate(formula: string): EngineValue | EngineValue[][];
  /** Last used row and column; ranges are clipped to them. */
  maxRow: number;
  maxCol: number;
}

/** Formula evaluations allowed per sheet before rules are skipped. */
const FORMULA_BUDGET = 200_000;

const isNum = (v: EngineValue): v is number => typeof v === 'number';
const isError = (v: EngineValue) => !!v && typeof v === 'object';
const blank = (v: EngineValue) => v === null || v === '';

function scalar(v: EngineValue | EngineValue[][]): EngineValue {
  return Array.isArray(v) ? (v[0]?.[0] ?? null) : v;
}

/** Excel's ordering for comparisons: numbers < text < logicals; text ignores case. */
function compare(a: EngineValue, b: EngineValue): number | undefined {
  if (isError(a) || isError(b)) return undefined;
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (v: EngineValue) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(a) - Number(b);
}

function truthy(v: EngineValue): boolean {
  return v === true || (typeof v === 'number' && v !== 0);
}

function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return NaN;
  const pos = (sorted.length - 1) * Math.min(1, Math.max(0, p));
  const lo = Math.floor(pos);
  return sorted[lo] + (sorted[Math.min(sorted.length - 1, lo + 1)] - sorted[lo]) * (pos - lo);
}

function mix(a: string, b: string, t: number): string {
  const ch = (hex: string, i: number) => parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16);
  const out = [0, 1, 2].map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t));
  return '#' + out.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

interface Acc {
  style?: DiffStyle;
  bar?: CfCell['bar'];
  icon?: CfCell['icon'];
  hideValue?: boolean;
  stopped?: boolean;
}

const STYLE_KEYS = ['fill', 'color', 'bold', 'italic', 'underline', 'strike'] as const;

/** Merges a lower-priority format under what higher-priority rules already set. */
function under(have: DiffStyle | undefined, add: DiffStyle): DiffStyle {
  if (!have) return add;
  const out: DiffStyle = { ...have };
  for (const k of STYLE_KEYS) if (out[k] === undefined && add[k] !== undefined) (out as Record<string, unknown>)[k] = add[k];
  if (add.border) out.border = { ...add.border, ...out.border };
  return out;
}

/**
 * What every rule on a sheet does to every cell, applied in priority order:
 * a format property set by a higher-priority rule is kept, and a matching
 * rule with "stop if true" ends evaluation for that cell.
 */
export function evaluateConditionalFormats(rules: CfRule[], sheet: number, ctx: CfContext): CfResult {
  const acc = new Map<number, Acc>();
  let budget = FORMULA_BUDGET;
  let skipped = 0;

  for (const rule of rules) {
    const cells: { row: number; col: number; v: EngineValue }[] = [];
    for (const r of rule.rects) {
      for (let row = r.r1; row <= Math.min(r.r2, ctx.maxRow); row++) {
        for (let col = r.c1; col <= Math.min(r.c2, ctx.maxCol); col++) cells.push({ row, col, v: ctx.value(row, col) });
      }
    }
    if (!cells.length) continue;
    const anchor = rule.rects[0];

    /** A rule formula's value at a cell; constants and absolute formulas are worked out once. */
    const formulaAt = (i: number): ((row: number, col: number) => EngineValue) | undefined => {
      const f = rule.formulas[i];
      if (f === undefined || f === '') return undefined;
      if (/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(f)) return () => Number(f);
      const quoted = /^"((?:[^"]|"")*)"$/.exec(f);
      if (quoted) return () => quoted[1].replace(/""/g, '"');
      let shift: (row: number, col: number) => string;
      try {
        shift = makeShifter(f, anchor.r1, anchor.c1);
      } catch {
        return () => ({ error: '#NAME?' });
      }
      if (shift(anchor.r1 + 1, anchor.c1 + 1) === f) {
        const once = scalar(ctx.calculate(f));
        return () => once;
      }
      return (row, col) => {
        budget--;
        return scalar(ctx.calculate(shift(row, col)));
      };
    };

    const apply = (i: number, what: Acc) => {
      const c = cells[i];
      const key = cellKey(c.row, c.col);
      let a = acc.get(key);
      if (!a) acc.set(key, (a = {}));
      if (a.stopped) return;
      if (what.style) a.style = under(a.style, what.style);
      a.bar ??= what.bar;
      a.icon ??= what.icon;
      a.hideValue ||= what.hideValue;
      if (rule.stopIfTrue) a.stopped = true;
    };
    const each = (test: (v: EngineValue, row: number, col: number) => boolean) => {
      cells.forEach((c, i) => {
        if (acc.get(cellKey(c.row, c.col))?.stopped) return;
        if (test(c.v, c.row, c.col)) apply(i, { style: rule.style });
      });
    };
    const needsFormulas = ['cellIs', 'expression', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod'];
    if (needsFormulas.includes(rule.type) && rule.formulas.length && cells.length > budget) {
      skipped++;
      continue;
    }

    const nums = cells.map((c) => c.v).filter(isNum);
    const sorted = [...nums].sort((a, b) => a - b);
    /** A threshold's value among the range's numbers. */
    const threshold = (t: Cfvo, fallback: 'min' | 'max'): number => {
      const min = sorted[0] ?? 0;
      const max = sorted[sorted.length - 1] ?? 0;
      switch (t.type) {
        case 'min':
          return min;
        case 'max':
          return max;
        case 'autoMin':
          return Math.min(0, min);
        case 'autoMax':
          return Math.max(0, max);
        case 'percent':
          return min + ((max - min) * Number(t.val ?? 0)) / 100;
        case 'percentile':
          return percentile(sorted, Number(t.val ?? 0) / 100);
        case 'num':
        case 'formula': {
          const v = t.val === undefined ? NaN : Number(t.val);
          if (!Number.isNaN(v)) return v;
          const r = scalar(ctx.calculate(t.val ?? ''));
          return typeof r === 'number' ? r : fallback === 'min' ? min : max;
        }
      }
    };

    switch (rule.type) {
      case 'cellIs': {
        const f1 = formulaAt(0);
        const f2 = formulaAt(1);
        if (!f1) break;
        each((v, row, col) => {
          if (isError(v)) return false;
          const a = compare(v, f1(row, col));
          if (a === undefined) return false;
          const b = f2 ? compare(v, f2(row, col)) : undefined;
          switch (rule.operator) {
            case 'between':
              return b !== undefined && a >= 0 && b <= 0;
            case 'notBetween':
              return b !== undefined && (a < 0 || b > 0);
            case 'equal':
              return a === 0;
            case 'notEqual':
              return a !== 0;
            case 'greaterThan':
              return a > 0;
            case 'lessThan':
              return a < 0;
            case 'greaterThanOrEqual':
              return a >= 0;
            case 'lessThanOrEqual':
              return a <= 0;
          }
          return false;
        });
        break;
      }
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith':
      case 'containsBlanks':
      case 'notContainsBlanks':
      case 'containsErrors':
      case 'notContainsErrors':
      case 'timePeriod':
      case 'expression': {
        // Excel writes an equivalent formula for every one of these.
        const f = formulaAt(0);
        if (f) {
          each((_v, row, col) => truthy(f(row, col)));
          break;
        }
        const text = (rule.text ?? '').toLowerCase();
        const str = (v: EngineValue) => (v === null || isError(v) ? '' : String(v).toLowerCase());
        const tests: Record<string, (v: EngineValue) => boolean> = {
          containsText: (v) => str(v).includes(text),
          notContainsText: (v) => !str(v).includes(text),
          beginsWith: (v) => str(v).startsWith(text),
          endsWith: (v) => str(v).endsWith(text),
          containsBlanks: (v) => blank(v) || (typeof v === 'string' && !v.trim()),
          notContainsBlanks: (v) => !(blank(v) || (typeof v === 'string' && !v.trim())),
          containsErrors: (v) => isError(v),
          notContainsErrors: (v) => !isError(v),
        };
        const test = tests[rule.type];
        if (test) each(test);
        break;
      }
      case 'top10': {
        if (!sorted.length) break;
        const n = rule.percent ? Math.max(1, Math.floor((sorted.length * (rule.rank ?? 10)) / 100)) : (rule.rank ?? 10);
        const k = Math.min(sorted.length, n) - 1;
        const cut = rule.bottom ? sorted[k] : sorted[sorted.length - 1 - k];
        each((v) => isNum(v) && (rule.bottom ? v <= cut : v >= cut));
        break;
      }
      case 'aboveAverage': {
        if (!nums.length) break;
        const mean = nums.reduce((s, x) => s + x, 0) / nums.length;
        const sd = Math.sqrt(nums.reduce((s, x) => s + (x - mean) ** 2, 0) / nums.length);
        const line = mean + (rule.stdDev ? (rule.aboveAverage ? 1 : -1) * rule.stdDev * sd : 0);
        each((v) => isNum(v) && (rule.aboveAverage ? v > line || (!!rule.equalAverage && v === line) : v < line || (!!rule.equalAverage && v === line)));
        break;
      }
      case 'duplicateValues':
      case 'uniqueValues': {
        const norm = (v: EngineValue) => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v === 'number' ? 'n' + v : typeof v === 'boolean' ? 'b' + v : undefined);
        const counts = new Map<string, number>();
        for (const c of cells) {
          const k = norm(c.v);
          if (k !== undefined && !blank(c.v)) counts.set(k, (counts.get(k) ?? 0) + 1);
        }
        each((v) => {
          const k = norm(v);
          if (k === undefined || blank(v)) return false;
          const n = counts.get(k) ?? 0;
          return rule.type === 'duplicateValues' ? n > 1 : n === 1;
        });
        break;
      }
      case 'colorScale': {
        const s = rule.colorScale;
        if (!s || s.colors.length < 2 || !sorted.length) break;
        const stops = s.cfvo.map((t, i) => threshold(t, i === 0 ? 'min' : 'max'));
        cells.forEach((c, i) => {
          if (!isNum(c.v)) return;
          // Three-colour scales blend low→mid below the midpoint and mid→high above it.
          const upper = s.colors.length > 2 && c.v > stops[1] ? 1 : 0;
          const lo = stops[upper];
          const hi = stops[upper + 1];
          const t = hi === lo ? 1 : (Math.min(hi, Math.max(lo, c.v)) - lo) / (hi - lo);
          apply(i, { style: { fill: mix(s.colors[upper], s.colors[upper + 1], t) } });
        });
        break;
      }
      case 'dataBar': {
        const b = rule.dataBar;
        if (!b || !sorted.length) break;
        const lo = threshold(b.cfvo[0] ?? { type: 'min', gte: true }, 'min');
        const hi = threshold(b.cfvo[1] ?? { type: 'max', gte: true }, 'max');
        cells.forEach((c, i) => {
          if (!isNum(c.v)) return;
          const t = hi === lo ? 1 : (Math.min(hi, Math.max(lo, c.v)) - lo) / (hi - lo);
          const len = (b.minLength + t * (b.maxLength - b.minLength)) / 100;
          const color = c.v < 0 && b.negativeColor ? b.negativeColor : b.color;
          apply(i, { bar: { from: 0, to: len, color, gradient: b.gradient }, hideValue: !b.showValue });
        });
        break;
      }
      case 'iconSet': {
        const s = rule.iconSet;
        if (!s || !sorted.length) break;
        const cuts = s.cfvo.map((t) => ({ at: threshold(t, 'min'), gte: t.gte }));
        cells.forEach((c, i) => {
          if (!isNum(c.v)) return;
          let index = 0;
          for (let k = 1; k < cuts.length; k++) if (cuts[k].gte ? c.v >= cuts[k].at : c.v > cuts[k].at) index = k;
          if (s.reverse) index = cuts.length - 1 - index;
          apply(i, { icon: { set: s.set, index }, hideValue: !s.showValue });
        });
        break;
      }
    }
  }

  // Share identical formats between cells.
  const styles: DiffStyle[] = [];
  const styleIds = new Map<string, number>();
  const cells: CfCell[] = [];
  for (const [key, a] of acc) {
    if (!a.style && !a.bar && !a.icon) continue;
    const cell: CfCell = { row: Math.floor(key / MAX_COLS), col: key % MAX_COLS };
    if (a.style) {
      const k = JSON.stringify(a.style);
      let id = styleIds.get(k);
      if (id === undefined) {
        id = styles.push(a.style) - 1;
        styleIds.set(k, id);
      }
      cell.style = id;
    }
    if (a.bar) cell.bar = a.bar;
    if (a.icon) cell.icon = a.icon;
    if (a.hideValue) cell.hideValue = true;
    cells.push(cell);
  }
  return { sheet, styles, cells, skipped };
}

// ---------------------------------------------------------------------------
// Writing rules

/** Icon sets Excel 2007 knows; others are Excel 2010's, kept in the sheet's extensions. */
const ICON_SETS_2007 = new Set([
  '3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
  '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters',
]);

const escapeText = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const quoteAttr = (s: string) => escapeText(s).replace(/"/g, '&quot;');
const argb = (hex: string) => `FF${hex.replace(/^#/, '').toUpperCase()}`;

/** A threshold as Excel 2007 writes one: Excel 2010's automatic ends are its minimum and maximum. */
function cfvoXml(p: string, c: Cfvo, icons: boolean): string {
  const type = c.type === 'autoMin' ? 'min' : c.type === 'autoMax' ? 'max' : c.type;
  return `<${p}cfvo type="${type}"${c.val !== undefined && type !== 'min' && type !== 'max' ? ` val="${quoteAttr(c.val)}"` : ''}${icons && !c.gte ? ' gte="0"' : ''}/>`;
}

/** One rule as Excel writes it in a worksheet's <conditionalFormatting>, with its place in the order. */
function cfRuleXml(p: string, r: CfRule, priority: number): string {
  const attrs = [`type="${r.type}"`];
  if (r.dxfId !== undefined) attrs.push(`dxfId="${r.dxfId}"`);
  attrs.push(`priority="${priority}"`);
  if (r.stopIfTrue) attrs.push('stopIfTrue="1"');
  if (r.type === 'aboveAverage') {
    if (r.aboveAverage === false) attrs.push('aboveAverage="0"');
    if (r.equalAverage) attrs.push('equalAverage="1"');
    if (r.stdDev !== undefined) attrs.push(`stdDev="${r.stdDev}"`);
  }
  if (r.rank !== undefined) attrs.push(`rank="${r.rank}"`);
  if (r.percent) attrs.push('percent="1"');
  if (r.bottom) attrs.push('bottom="1"');
  if (r.operator) attrs.push(`operator="${r.operator}"`);
  if (r.text !== undefined) attrs.push(`text="${quoteAttr(r.text)}"`);
  let body = r.formulas.map((f) => `<${p}formula>${escapeText(f)}</${p}formula>`).join('');
  if (r.colorScale) {
    const { cfvo, colors } = r.colorScale;
    body += `<${p}colorScale>${cfvo.map((c) => cfvoXml(p, c, false)).join('')}${colors.map((c) => `<${p}color rgb="${argb(c)}"/>`).join('')}</${p}colorScale>`;
  }
  if (r.dataBar) {
    const b = r.dataBar;
    const lengths = (b.minLength !== 10 ? ` minLength="${b.minLength}"` : '') + (b.maxLength !== 90 ? ` maxLength="${b.maxLength}"` : '');
    body += `<${p}dataBar${lengths}${b.showValue ? '' : ' showValue="0"'}>${b.cfvo.map((c) => cfvoXml(p, c, false)).join('')}<${p}color rgb="${argb(b.color)}"/></${p}dataBar>`;
  }
  if (r.iconSet) {
    const i = r.iconSet;
    const set = ICON_SETS_2007.has(i.set) ? i.set : '3TrafficLights1';
    const named = set === '3TrafficLights1' ? '' : ` iconSet="${set}"`;
    body += `<${p}iconSet${named}${i.reverse ? ' reverse="1"' : ''}${i.showValue ? '' : ' showValue="0"'}>${i.cfvo.map((c) => cfvoXml(p, c, true)).join('')}</${p}iconSet>`;
  }
  if (r.extId) {
    body +=
      `<${p}extLst><${p}ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
      `<x14:id>${escapeText(r.extId)}</x14:id></${p}ext></${p}extLst>`;
  }
  return `<${p}cfRule ${attrs.join(' ')}>${body}</${p}cfRule>`;
}

/**
 * The worksheet with its conditional formatting as `rules` have it, in
 * their order: each rule in a <conditionalFormatting> of its own, in its
 * place among the worksheet's elements. Excel 2010's part of a rule, in the
 * sheet's extensions, follows its rule's ranges and place, and goes with it.
 */
export function writeConditionalFormats(xml: string, p: string, rules: readonly CfRule[]): string {
  const priority = new Map(rules.map((r, i) => [r, i + 1]));
  const blocks = rules
    .filter((r) => !r.x14Only && r.rects.length)
    .map((r) => `<${p}conditionalFormatting sqref="${r.rects.map(rangeToA1).join(' ')}">${cfRuleXml(p, r, priority.get(r)!)}</${p}conditionalFormatting>`);
  xml = placeSheetElement(xml, 'conditionalFormatting', blocks.join(''));
  const byId = new Map(rules.filter((r) => r.extId && r.rects.length).map((r) => [r.extId!, r]));
  const ext = /<(\w+:)?ext\b[^>]*\buri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"[^>]*>[\s\S]*?<\/\1?ext>/.exec(xml);
  if (!ext) return xml;
  const patched = ext[0].replace(/<(\w+:)conditionalFormatting\b([^>]*)>([\s\S]*?)<\/\1conditionalFormatting>/g, (_all, q: string, attrs: string, body: string) => {
    let ranges: Rect[] | undefined;
    const kept = body.replace(/<(\w+:)cfRule\b([^>]*)>([\s\S]*?)<\/\1cfRule>/g, (rule, _q: string, ruleAttrs: string) => {
      const r = byId.get(/\bid="([^"]*)"/.exec(ruleAttrs)?.[1] ?? '');
      if (!r) return '';
      ranges = r.rects;
      return r.x14Only ? rule.replace(/\bpriority="\d+"/, `priority="${priority.get(r)}"`) : rule;
    });
    if (!ranges) return '';
    const sqref = kept.replace(/(<(\w+:)?sqref>)[^<]*(<\/(\w+:)?sqref>)/, (_m, open: string, _a: string, close: string) => `${open}${ranges!.map(rangeToA1).join(' ')}${close}`);
    return `<${q}conditionalFormatting${attrs}>${sqref}</${q}conditionalFormatting>`;
  });
  // An extension left with no rules goes, and the extensions too when none is left.
  const empty = !/conditionalFormatting\b/.test(patched);
  xml = xml.slice(0, ext.index) + (empty ? '' : patched) + xml.slice(ext.index + ext[0].length);
  return xml.replace(/<(\w+:)?extLst>\s*<\/\1?extLst>/, '');
}

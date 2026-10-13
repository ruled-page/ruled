/**
 * Pivot tables worked out from their source, as Excel's Refresh works them
 * out: the definition (xl/pivotTables/pivotTableN.xml) and its cache's
 * fields (xl/pivotCache/pivotCacheDefinitionN.xml) read into a model, and
 * the model laid out over the source's rows in Excel's compact, outline or
 * tabular form, with subtotals, grand totals, filters and several values.
 */
import type { EngineValue } from '@ruled/formula/engine';
import { parseRange } from '@ruled/formula/refs';
import type { Rect } from './types';
import { attr, child, children, type XmlNode } from './xml';

export type PivotFn = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';

/** A value in a pivot cache, or a source cell: dates are their serial numbers. */
export type PivotValue = string | number | boolean | null | { error: string };

export interface PivotItem {
  /** The cache value it stands for; undefined for Excel's subtotal and grand total entries. */
  value?: PivotValue;
  hidden: boolean;
  /** Its details are hidden, as Excel's collapse button hides them. */
  collapsed: boolean;
  /** A name typed over the item's own. */
  name?: string;
  /** Its value's index among the cache's shared items; undefined for one the cache doesn't have yet. */
  x?: number;
}

export interface PivotField {
  /** The cache field's name, or the caption typed over it. */
  name: string;
  axis?: 'row' | 'col' | 'page';
  items: PivotItem[];
  compact: boolean;
  outline: boolean;
  /** Subtotals shown on the item's own row rather than below its details. */
  subtotalTop: boolean;
  /** Its items are subtotalled (Excel's automatic subtotal). */
  subtotals: boolean;
  sort: 'manual' | 'ascending' | 'descending';
  /** Items with no data are shown too. */
  showAll: boolean;
  /** The number format its items are shown in: the cache field's. */
  numFmtId?: number;
}

/**
 * How a value is shown, as Excel's Show Values As offers: as it is, or
 * against a total, another item, the items before it, or its siblings. The
 * first eight are the file format's own; the rest came with Excel 2010.
 */
export type PivotShowAs =
  | 'normal'
  | 'difference'
  | 'percent'
  | 'percentDiff'
  | 'runTotal'
  | 'percentOfRow'
  | 'percentOfCol'
  | 'percentOfTotal'
  | 'index'
  | 'percentOfParentRow'
  | 'percentOfParentCol'
  | 'percentOfParent'
  | 'percentOfRunningTotal'
  | 'rankAscending'
  | 'rankDescending';

export const SHOW_AS: PivotShowAs[] = [
  'normal',
  'percentOfTotal',
  'percentOfCol',
  'percentOfRow',
  'percent',
  'percentOfParentRow',
  'percentOfParentCol',
  'percentOfParent',
  'difference',
  'percentDiff',
  'runTotal',
  'percentOfRunningTotal',
  'rankAscending',
  'rankDescending',
  'index',
];

/** Show Values As kinds measured against a base field, and those against one of its items too. */
export const NEEDS_BASE_FIELD = new Set<PivotShowAs>(['difference', 'percent', 'percentDiff', 'runTotal', 'percentOfParent', 'percentOfRunningTotal', 'rankAscending', 'rankDescending']);
export const NEEDS_BASE_ITEM = new Set<PivotShowAs>(['difference', 'percent', 'percentDiff']);
/** Kinds Excel shows as percentages. */
export const PERCENTS = new Set<PivotShowAs>(['percent', 'percentDiff', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'percentOfParentRow', 'percentOfParentCol', 'percentOfParent', 'percentOfRunningTotal']);
/** The kinds the file format's showDataAs holds; the others are in Excel 2010's extension. */
export const FILE_SHOW_AS = new Set<PivotShowAs>(['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index']);
/** Excel's base items "(previous)" and "(next)". */
export const PREVIOUS_ITEM = 1048828;
export const NEXT_ITEM = 1048829;

/**
 * The number format a value is shown in: Excel's percentage for the kinds of
 * Show Values As that are percentages, plain numbers for ranks and the index,
 * and otherwise the value's own.
 */
export function shownFormat(d: PivotDataField): number | undefined {
  if (d.showAs && PERCENTS.has(d.showAs)) return 10;
  if (d.showAs === 'index' || d.showAs === 'rankAscending' || d.showAs === 'rankDescending') return 0;
  return d.numFmtId;
}

export interface PivotDataField {
  name: string;
  field: number;
  fn: PivotFn;
  numFmtId?: number;
  /** How it is shown, when not as it is, and against which field and item (an index among its items, or the previous or next). */
  showAs?: PivotShowAs;
  baseField?: number;
  baseItem?: number;
}

export interface PivotModel {
  name: string;
  /** Where the table is, page fields aside, and where its headers and data start inside it. */
  location: { ref: Rect; firstHeaderRow: number; firstDataRow: number; firstDataCol: number };
  /** The cache's source: a range on a sheet, or a table's or name's. */
  source?: { sheet?: string; ref?: string; name?: string };
  fields: PivotField[];
  /** Fields down the rows and across the columns, outermost first; -2 is the Values field. */
  rows: number[];
  cols: number[];
  /** Report filters: the field, and the item index picked (undefined for all, or several). */
  pages: { field: number; item?: number }[];
  data: PivotDataField[];
  rowGrand: boolean;
  colGrand: boolean;
  captions: { data: string; grand: string; rowHeader: string; colHeader: string };
  /** Why it can't be worked out here (kept as Excel saved it): grouped or calculated fields, another kind of source… */
  unsupported?: string;
}

const flag = (node: XmlNode | undefined, name: string, dflt: boolean) => {
  const v = attr(node, name);
  return v === undefined ? dflt : v === '1' || v === 'true';
};
const int = (node: XmlNode | undefined, name: string) => {
  const v = attr(node, name);
  return v === undefined ? undefined : Number(v);
};

/** Days from Excel's 1900 epoch to an ISO date-time, as a cache's date items give them. */
function serialOf(iso: string): number {
  const t = Date.parse(iso.endsWith('Z') ? iso : `${iso}Z`);
  return Number.isNaN(t) ? 0 : t / 86_400_000 + 25_569;
}

/** A cache's shared item, or a record's value, as a value. */
export function sharedValue(n: XmlNode): PivotValue {
  const v = attr(n, 'v');
  switch (n.name) {
    case 'n':
      return Number(v);
    case 'b':
      return v === '1' || v === 'true';
    case 'd':
      return serialOf(v ?? '');
    case 'e':
      return { error: v ?? '#N/A' };
    case 'm':
      return null;
    default:
      return v ?? '';
  }
}

const FNS = new Set<PivotFn>(['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp']);

/** A pivot table's definition and its cache's definition read into a model. */
export function readPivot(table: XmlNode, cache: XmlNode | undefined): PivotModel {
  const loc = child(table, 'location');
  const cacheFields = children(child(cache, 'cacheFields'), 'cacheField');
  const source = child(child(cache, 'cacheSource'), 'worksheetSource');
  let unsupported: string | undefined;
  if (!cache) unsupported = 'cache';
  else if (attr(child(cache, 'cacheSource'), 'type') !== 'worksheet' || !source) unsupported = 'source';
  else if (cacheFields.some((f) => child(f, 'fieldGroup') || attr(f, 'formula') !== undefined || flag(f, 'databaseField', true) === false)) unsupported = 'grouped';
  const tableCompact = flag(table, 'compact', true);
  const fields = children(child(table, 'pivotFields'), 'pivotField').map((pf, i): PivotField => {
    const cf = cacheFields[i];
    const shared = child(cf, 'sharedItems')?.children.map(sharedValue) ?? [];
    const axis = attr(pf, 'axis');
    if (children(child(pf, 'items'), 'item').some((it) => attr(it, 'f') === '1' || attr(it, 't') === 'blank')) unsupported ??= 'calculatedItem';
    return {
      name: attr(pf, 'name') ?? attr(cf, 'name') ?? `Field${i + 1}`,
      axis: axis === 'axisRow' ? 'row' : axis === 'axisCol' ? 'col' : axis === 'axisPage' ? 'page' : undefined,
      items: children(child(pf, 'items'), 'item')
        .filter((it) => (attr(it, 't') ?? 'data') === 'data')
        .map((it) => {
          const x = int(it, 'x');
          return {
            value: x === undefined ? undefined : shared[x],
            ...(x !== undefined && { x }),
            hidden: flag(it, 'h', false),
            collapsed: !flag(it, 'sd', true),
            ...(attr(it, 'n') !== undefined && { name: attr(it, 'n') }),
          };
        }),
      compact: flag(pf, 'compact', tableCompact),
      outline: flag(pf, 'outline', true),
      subtotalTop: flag(pf, 'subtotalTop', true),
      subtotals: flag(pf, 'defaultSubtotal', true),
      sort: (attr(pf, 'sortType') as PivotField['sort']) ?? 'manual',
      showAll: flag(pf, 'showAll', true),
      numFmtId: int(cf, 'numFmtId'),
    };
  });
  const data = children(child(table, 'dataFields'), 'dataField').map((d): PivotDataField => {
    const fn = (attr(d, 'subtotal') ?? 'sum') as PivotFn;
    if (!FNS.has(fn)) unsupported ??= 'function';
    // Excel 2010's kinds are in an extension, beside the file format's own.
    const later = children(child(d, 'extLst'), 'ext').map((e) => attr(child(e, 'dataField'), 'pivotShowAs')).find((v) => v !== undefined);
    const showAs = (later ?? attr(d, 'showDataAs') ?? 'normal') as PivotShowAs;
    if (!SHOW_AS.includes(showAs)) unsupported ??= 'showDataAs';
    const field = int(d, 'fld') ?? 0;
    return {
      name: attr(d, 'name') ?? `${fn} of ${fields[field]?.name ?? ''}`,
      field,
      fn,
      numFmtId: int(d, 'numFmtId'),
      ...(showAs !== 'normal' && { showAs, baseField: int(d, 'baseField') ?? 0, baseItem: int(d, 'baseItem') ?? 0 }),
    };
  });
  const fieldList = (name: string) => children(child(table, name), 'field').map((f) => int(f, 'x') ?? 0);
  const range = parseRange(attr(loc, 'ref') ?? '') ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
  return {
    name: attr(table, 'name') ?? '',
    location: { ref: range, firstHeaderRow: int(loc, 'firstHeaderRow') ?? 1, firstDataRow: int(loc, 'firstDataRow') ?? 1, firstDataCol: int(loc, 'firstDataCol') ?? 1 },
    ...(source && { source: { sheet: attr(source, 'sheet'), ref: attr(source, 'ref'), name: attr(source, 'name') } }),
    fields,
    rows: fieldList('rowFields'),
    cols: fieldList('colFields'),
    pages: children(child(table, 'pageFields'), 'pageField').map((p) => ({ field: int(p, 'fld') ?? 0, item: int(p, 'item') })),
    data,
    rowGrand: flag(table, 'rowGrandTotals', true),
    colGrand: flag(table, 'colGrandTotals', true),
    captions: {
      data: attr(table, 'dataCaption') ?? 'Values',
      grand: attr(table, 'grandTotalCaption') ?? 'Grand Total',
      rowHeader: attr(table, 'rowHeaderCaption') ?? 'Row Labels',
      colHeader: attr(table, 'colHeaderCaption') ?? 'Column Labels',
    },
    ...(unsupported && { unsupported }),
  };
}

/** A pivot table as Excel's Insert PivotTable makes it: fields from its source's headings, none of them placed yet. */
export function newPivot(name: string, fields: { name: string; numFmtId?: number }[], source: NonNullable<PivotModel['source']>, ref: Rect): PivotModel {
  return {
    name,
    location: { ref, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 0 },
    source,
    fields: fields.map((f) => ({ name: f.name, items: [], compact: true, outline: true, subtotalTop: true, subtotals: true, sort: 'manual', showAll: false, numFmtId: f.numFmtId })),
    rows: [],
    cols: [],
    pages: [],
    data: [],
    rowGrand: true,
    colGrand: true,
    captions: { data: 'Values', grand: 'Grand Total', rowHeader: 'Row Labels', colHeader: 'Column Labels' },
  };
}

// ---------------------------------------------------------------------------
// Changing a pivot table's fields

/**
 * What Excel's PivotTable Fields pane sets: the fields in each area,
 * outermost first (-2 being the Values field among the rows or columns),
 * the item each report filter picks, how each value is summarised, and the
 * items each field hides. Items are named by their keys (`keyOf`).
 */
export interface PivotSetup {
  rows: number[];
  cols: number[];
  pages: { field: number; item?: string }[];
  data: PivotSetupValue[];
  hidden: Record<number, string[]>;
}

/** A value in the pane: its field, summary and name, and how it is shown, against a base item named by its key, or 'previous' or 'next'. */
export interface PivotSetupValue {
  field: number;
  fn: PivotFn;
  name: string;
  showAs?: PivotShowAs;
  baseField?: number;
  baseItem?: string;
}

/** A pivot table's fields as the pane shows them. */
export function setupOf(model: PivotModel): PivotSetup {
  const hidden: Record<number, string[]> = {};
  model.fields.forEach((f, i) => {
    const keys = f.items.filter((it) => it.hidden && it.value !== undefined).map((it) => keyOf(it.value));
    if (keys.length) hidden[i] = keys;
  });
  return {
    rows: [...model.rows],
    cols: [...model.cols],
    pages: model.pages.map((p) => {
      const value = p.item !== undefined ? model.fields[p.field]?.items[p.item]?.value : undefined;
      return value !== undefined ? { field: p.field, item: keyOf(value) } : { field: p.field };
    }),
    data: model.data.map((d) => {
      if (!d.showAs || d.showAs === 'normal') return { field: d.field, fn: d.fn, name: d.name };
      const bf = d.baseField ?? 0;
      const item = d.baseItem === PREVIOUS_ITEM ? 'previous' : d.baseItem === NEXT_ITEM ? 'next' : keyOf(model.fields[bf]?.items[d.baseItem ?? 0]?.value);
      return { field: d.field, fn: d.fn, name: d.name, showAs: d.showAs, baseField: bf, baseItem: item };
    }),
    hidden,
  };
}

/**
 * The pivot table with its fields set as `setup` says. A field put on the
 * rows, columns or filters for the first time has its items from `records`
 * (the source's rows), sorted as Excel first lists them; one already there
 * keeps its order, with values new since after it.
 */
export function withSetup(model: PivotModel, setup: PivotSetup, records: Rec[]): PivotModel {
  const n = model.fields.length;
  const valid = (x: number) => Number.isInteger(x) && x >= 0 && x < n;
  const data = setup.data.filter((d) => valid(d.field) && FNS.has(d.fn));
  // The Values field is among the rows or columns only with several values, and across the columns unless put down the rows.
  const axis = (list: number[]) => [...new Set(list.filter((x) => valid(x) || (x === -2 && data.length > 1)))];
  let rows = axis(setup.rows);
  let cols = axis(setup.cols).filter((x) => !rows.includes(x));
  if (data.length > 1 && !rows.includes(-2) && !cols.includes(-2)) cols = [...cols, -2];
  const pages = setup.pages.filter((p, i) => valid(p.field) && !rows.includes(p.field) && !cols.includes(p.field) && setup.pages.findIndex((q) => q.field === p.field) === i);
  const fields = model.fields.map((f, i): PivotField => {
    const on = rows.includes(i) ? 'row' : cols.includes(i) ? 'col' : pages.some((p) => p.field === i) ? 'page' : undefined;
    const hidden = new Set(setup.hidden[i] ?? []);
    const known = f.items.filter((it) => it.value !== undefined);
    const seen = new Set(known.map((it) => keyOf(it.value)));
    const fresh: PivotItem[] = [];
    if (on) {
      for (const r of records) {
        const k = keyOf(r[i]);
        if (seen.has(k)) continue;
        seen.add(k);
        fresh.push({ value: r[i] ?? null, hidden: false, collapsed: false });
      }
      if (!known.length) fresh.sort((a, b) => compareValues(a.value, b.value));
    }
    return { ...f, axis: on, items: [...known, ...fresh].map((it) => ({ ...it, hidden: !!on && hidden.has(keyOf(it.value)) })) };
  });
  return {
    ...model,
    fields,
    rows,
    cols,
    pages: pages.map((p) => {
      const at = p.item === undefined ? -1 : fields[p.field].items.findIndex((it) => keyOf(it.value) === p.item);
      return at < 0 ? { field: p.field } : { field: p.field, item: at };
    }),
    // A value keeps its field's format when summarised another way, and a new one takes its source's, as Excel 365 does; counts are plain numbers.
    data: data.map((d): PivotDataField => {
      const counted = d.fn === 'count' || d.fn === 'countNums';
      const was = model.data.find((x) => x.field === d.field);
      const value: PivotDataField = { name: d.name, field: d.field, fn: d.fn, numFmtId: counted ? undefined : was ? was.numFmtId : model.fields[d.field].numFmtId || undefined };
      if (!d.showAs || d.showAs === 'normal' || !SHOW_AS.includes(d.showAs)) return value;
      const bf = valid(d.baseField ?? -1) ? d.baseField! : 0;
      const at = fields[bf].items.findIndex((it) => keyOf(it.value) === d.baseItem);
      return { ...value, showAs: d.showAs, baseField: bf, baseItem: d.baseItem === 'previous' ? PREVIOUS_ITEM : d.baseItem === 'next' ? NEXT_ITEM : Math.max(0, at) };
    }),
  };
}

// ---------------------------------------------------------------------------
// Laying a pivot table out

/** A cell of a laid-out pivot table: where it is from the table's top-left, what it holds, and what kind of cell it is. */
export interface PivotCell {
  row: number;
  col: number;
  value: EngineValue;
  kind: 'header' | 'label' | 'value' | 'total' | 'grand' | 'page';
  /** The number format a value is shown in: its data field's. */
  numFmtId?: number;
}

export interface PivotLayout {
  /** Rows and columns the table covers, page fields aside. */
  height: number;
  width: number;
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
  /** Cells inside the table, and the report filters' above it (negative rows). */
  cells: PivotCell[];
}

type Rec = PivotValue[];

/** Which items match each other: Excel's are case-insensitive, and blank is its own item. */
export function keyOf(v: PivotValue | undefined): string {
  if (v === null || v === undefined || v === '') return 'm';
  if (typeof v === 'string') return `s${v.toLowerCase()}`;
  if (typeof v === 'number') return `n${v}`;
  if (typeof v === 'boolean') return `b${v}`;
  return `e${v.error}`;
}

/** A level of the rows or columns: a field, or the Values field (the data fields as its items). */
type Level = { field: number } | { values: true };

interface Entry {
  /** Labels in the row label columns (rows) or the header rows (columns), by position. */
  labels: Map<number, PivotValue>;
  /** The records it sums. */
  recs: Rec[];
  kind: 'item' | 'total' | 'grand';
  /** The data field it shows, when the Values field is on its axis. */
  data?: number;
  /** A parent item's own row, which shows its subtotal only when the field has one on top. */
  blank?: boolean;
  /** The items it is under, field by field from the outermost, its own last: none for a grand total. */
  path: [number, string][];
}

function compareValues(a: PivotValue | undefined, b: PivotValue | undefined): number {
  const rank = (v: PivotValue | undefined) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null || v === undefined ? 4 : 3);
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { sensitivity: 'base' });
}

/** A data field's summary of the records, as Excel's: Sum, Count, Average and the rest. */
export function summarize(fn: PivotFn, values: PivotValue[]): EngineValue {
  const error = values.find((v): v is { error: string } => v !== null && typeof v === 'object');
  if (error && fn !== 'count') return { error: error.error };
  const nums = values.filter((v): v is number => typeof v === 'number');
  const n = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  const mean = n ? sum / n : 0;
  const sq = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  switch (fn) {
    case 'sum':
      return sum;
    case 'count':
      return values.filter((v) => v !== null && v !== '').length;
    case 'countNums':
      return n;
    case 'average':
      return n ? mean : { error: '#DIV/0!' };
    case 'max':
      return n ? Math.max(...nums) : 0;
    case 'min':
      return n ? Math.min(...nums) : 0;
    case 'product':
      return n ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'stdDev':
      return n > 1 ? Math.sqrt(sq / (n - 1)) : { error: '#DIV/0!' };
    case 'stdDevp':
      return n ? Math.sqrt(sq / n) : { error: '#DIV/0!' };
    case 'var':
      return n > 1 ? sq / (n - 1) : { error: '#DIV/0!' };
    case 'varp':
      return n ? sq / n : { error: '#DIV/0!' };
  }
}

/**
 * The pivot table laid out over `records` (one per source row, a value per
 * cache field), as Excel lays it out on refresh. `label` gives an item's
 * text in its field's number format.
 */
export function layoutPivot(model: PivotModel, records: Rec[], label: (v: PivotValue, numFmtId: number | undefined) => string): PivotLayout {
  const { fields, data } = model;
  const blankText = '(blank)';
  const itemText = (f: number, v: PivotValue | undefined, name?: string) => name ?? (v === null || v === undefined || v === '' ? blankText : label(v, fields[f]?.numFmtId));

  // Records the report filters and hidden items let through.
  const hiddenKeys = fields.map((f) => new Set(f.items.filter((i) => i.hidden).map((i) => keyOf(i.value))));
  const keep = records.filter((rec) => {
    for (const p of model.pages) {
      const pick = p.item !== undefined ? fields[p.field]?.items[p.item] : undefined;
      if (pick && keyOf(rec[p.field]) !== keyOf(pick.value)) return false;
    }
    return fields.every((f, i) => (f.axis ? !hiddenKeys[i].has(keyOf(rec[i])) : true));
  });

  /** A field's items in the order shown: Excel's own order, then values new since, sorted if the field is. */
  const itemsOf = (f: number, recs: Rec[]): { value: PivotValue; name?: string; collapsed: boolean }[] => {
    const field = fields[f];
    const known = new Map<string, { value: PivotValue; name?: string; collapsed: boolean }>();
    for (const it of field.items) if (!it.hidden && it.value !== undefined) known.set(keyOf(it.value), { value: it.value, name: it.name, collapsed: it.collapsed });
    const present = new Map<string, PivotValue>();
    for (const r of recs) if (!present.has(keyOf(r[f]))) present.set(keyOf(r[f]), r[f]);
    const out = [...known].filter(([k]) => field.showAll || present.has(k)).map(([, v]) => v);
    for (const [k, v] of present) if (!known.has(k) && !hiddenKeys[f].has(k)) out.push({ value: v, collapsed: false });
    if (field.sort !== 'manual') out.sort((a, b) => (field.sort === 'descending' ? -1 : 1) * compareValues(a.value, b.value));
    return out;
  };

  const values = data.length > 1;
  const level = (x: number): Level => (x === -2 ? { values: true } : { field: x });
  const rowLevels = model.rows.filter((x) => x !== -2 || values).map(level);
  const colLevels = model.cols.filter((x) => x !== -2 || values).map(level);
  // With several data fields and nowhere said, they go across the columns, innermost.
  if (values && !rowLevels.some((l) => 'values' in l) && !colLevels.some((l) => 'values' in l)) colLevels.push({ values: true });
  const valuesOnRows = rowLevels.some((l) => 'values' in l);

  // The form of each row field: compact, outline or tabular, as its settings say.
  const formOf = (l: Level, i: number): 'compact' | 'outline' | 'tabular' => {
    const f = 'field' in l ? fields[l.field] : i > 0 && 'field' in rowLevels[i - 1] ? fields[(rowLevels[i - 1] as { field: number }).field] : undefined;
    if (!f) return 'compact';
    return !f.outline ? 'tabular' : f.compact ? 'compact' : 'outline';
  };
  // Which label column each row level is in: compact levels share their parent's.
  const rowCol: number[] = [];
  rowLevels.forEach((l, i) => rowCol.push(i === 0 ? 0 : formOf(l, i) === 'compact' && formOf(rowLevels[i - 1], i - 1) === 'compact' ? rowCol[i - 1] : rowCol[i - 1] + 1));
  const onlyValuesAcross = colLevels.length === 1 && 'values' in colLevels[0];
  // With no row fields there is a label column only for the column fields' labels.
  const labelCols = rowLevels.length ? rowCol[rowCol.length - 1] + 1 : colLevels.length && !onlyValuesAcross ? 1 : 0;

  /** The rows (or columns) under one level, recursively. */
  const entries = (levels: Level[], at: number, recs: Rec[], axis: 'row' | 'col', path: [number, string][] = []): Entry[] => {
    const l = levels[at];
    const last = at === levels.length - 1;
    const pos = axis === 'row' ? rowCol[at] : at;
    const items =
      'values' in l
        ? data.map((d, i) => ({ value: d.name as PivotValue, data: i, collapsed: false, name: undefined as string | undefined }))
        : itemsOf(l.field, recs).map((it) => ({ ...it, data: undefined as number | undefined }));
    const form = axis === 'row' ? formOf(l, at) : 'tabular';
    const field = 'field' in l ? fields[l.field] : undefined;
    const out: Entry[] = [];
    for (const it of items) {
      const mine = 'values' in l ? recs : recs.filter((r) => keyOf(r[l.field]) === keyOf(it.value));
      const text = 'values' in l ? String(it.value) : itemText(l.field, it.value, it.name);
      const withData = (e: Entry) => (it.data !== undefined ? { ...e, data: it.data } : e);
      const own: [number, string][] = 'values' in l ? path : [...path, [l.field, keyOf(it.value)]];
      if (last || it.collapsed) {
        out.push(withData({ labels: new Map([[pos, text]]), recs: mine, kind: 'item', path: own }));
        continue;
      }
      const inner = entries(levels, at + 1, mine, axis, own).map(withData);
      // A field is subtotalled when another field is under it, not just the data fields.
      const subtotal = (!!field?.subtotals && levels.slice(at + 1).some((x) => 'field' in x)) || 'values' in l;
      if (form === 'tabular') {
        // The item's label shares its first detail's row (or column).
        if (inner[0]) inner[0].labels.set(pos, text);
        out.push(...inner);
      } else {
        // Its own row shows its subtotal on top, unless the data fields come below it, each on its own row.
        const valuesBelow = levels.slice(at + 1).some((x) => 'values' in x);
        out.push(withData({ labels: new Map([[pos, text]]), recs: mine, kind: 'item', path: own, blank: !(subtotal && field?.subtotalTop !== false) || 'values' in l || valuesBelow }));
        out.push(...inner);
      }
      if (subtotal && !('values' in l) && (form === 'tabular' || field?.subtotalTop === false)) {
        if (valuesOnRows && axis === 'row' && levels.slice(at + 1).some((x) => 'values' in x)) {
          data.forEach((d, i) => out.push({ labels: new Map([[pos, `${text} ${d.name}`]]), recs: mine, kind: 'total', data: i, path: own }));
        } else {
          out.push(withData({ labels: new Map([[pos, `${text} Total`]]), recs: mine, kind: 'total', path: own }));
        }
      }
    }
    return out;
  };

  const whole = (): Entry => ({ labels: new Map<number, PivotValue>(), recs: keep, kind: 'item', path: [] });
  const rowsOut = rowLevels.length ? entries(rowLevels, 0, keep, 'row') : [whole()];
  const colsOut = colLevels.length ? entries(colLevels, 0, keep, 'col') : [whole()];
  // Grand totals: one per data field when the Values field is on that axis with others.
  const grand = (levels: Level[]): Entry[] => {
    const onAxis = levels.some((l) => 'values' in l);
    if (onAxis && levels.length > 1) return data.map((d, i) => ({ labels: new Map([[0, `Total ${d.name}`]]), recs: keep, kind: 'grand', data: i, path: [] }));
    if (onAxis) return [];
    return levels.length ? [{ labels: new Map([[0, model.captions.grand]]), recs: keep, kind: 'grand', path: [] }] : [];
  };
  if (model.rowGrand) rowsOut.push(...grand(rowLevels));
  if (model.colGrand) colsOut.push(...grand(colLevels));

  // Header rows: one per column level, the Values field's included, and one for the row labels' headers.
  const headerRows = colLevels.length ? colLevels.length + 1 : 1;
  const firstDataRow = headerRows;
  const cells: PivotCell[] = [];
  const put = (row: number, col: number, value: EngineValue, kind: PivotCell['kind'], numFmtId?: number) => {
    if (value === null || value === '') return;
    cells.push({ row, col, value, kind, ...(numFmtId !== undefined && { numFmtId }) });
  };

  const compactRows = rowLevels.length > 0 && formOf(rowLevels[0], 0) === 'compact';
  const lastHeader = headerRows - 1;
  // The row labels' headers: Excel's caption in compact form, the fields' names in the others.
  if (rowLevels.length) {
    if (compactRows) put(lastHeader, 0, model.captions.rowHeader, 'header');
    rowLevels.forEach((l, i) => {
      if (i > 0 && rowCol[i] === rowCol[i - 1]) return;
      if (i === 0 && compactRows) return;
      put(lastHeader, rowCol[i], 'values' in l ? model.captions.data : fields[l.field].name, 'header');
    });
  }
  if (colLevels.length) {
    // The caption row: the one data field's name, then the column fields' heading: Excel's "Values" when the Values field is first across.
    if (!values) put(0, 0, data[0]?.name ?? '', 'header');
    const first = colLevels[0];
    // Compact form captions it Excel's way; the others name the field. With no row fields, the column field's own form says which.
    const compact = rowLevels.length ? compactRows : 'field' in first && fields[first.field].compact && fields[first.field].outline;
    put(0, labelCols, 'values' in first ? model.captions.data : compact ? model.captions.colHeader : fields[first.field].name, 'header');
  }
  // Column labels, each on its level's row, then the data fields' names where only they go across.
  colsOut.forEach((c, i) => {
    for (const [lvl, text] of c.labels) put(lvl + 1, labelCols + i, text, c.kind === 'grand' ? 'grand' : c.kind === 'total' ? 'total' : 'header');
  });
  if (!colLevels.length && !valuesOnRows) put(0, labelCols, data[0]?.name ?? '', 'header');
  if (!colLevels.length && valuesOnRows && !rowLevels.length) put(0, labelCols, model.captions.data, 'header');

  /** A data field's summary of the records under the given items, as a cell elsewhere in the table shows it. */
  const memo = new Map<string, EngineValue | undefined>();
  const valueAt = (path: [number, string][], di: number): EngineValue | undefined => {
    const key = `${di}|${path.map(([f, k]) => `${f}=${k}`).sort().join('|')}`;
    if (memo.has(key)) return memo.get(key);
    const recs = keep.filter((rec) => path.every(([f, k]) => keyOf(rec[f]) === k));
    const value = recs.length ? summarize(data[di].fn, recs.map((rec) => rec[data[di].field])) : undefined;
    memo.set(key, value);
    return value;
  };
  const orders = new Map<number, string[]>();
  const orderOf = (f: number) => orders.get(f) ?? orders.set(f, itemsOf(f, keep).map((it) => keyOf(it.value))).get(f)!;
  const ratio = (a: number, b: EngineValue | undefined): EngineValue => (typeof b === 'number' && b !== 0 ? a / b : { error: '#DIV/0!' });

  /**
   * A value as Show Values As shows it, from the totals and the other cells it
   * is measured against; null where Excel leaves the cell empty, as where the
   * base field isn't among those the cell is under.
   */
  const showAs = (value: EngineValue, rowPath: [number, string][], colPath: [number, string][], di: number): EngineValue | null => {
    const d = data[di];
    if (typeof value !== 'number') return value;
    const all = [...rowPath, ...colPath];
    switch (d.showAs) {
      case 'percentOfTotal':
        return ratio(value, valueAt([], di));
      case 'percentOfRow':
        return ratio(value, valueAt(rowPath, di));
      case 'percentOfCol':
        return ratio(value, valueAt(colPath, di));
      case 'percentOfParentRow':
        return ratio(value, valueAt([...rowPath.slice(0, -1), ...colPath], di));
      case 'percentOfParentCol':
        return ratio(value, valueAt([...rowPath, ...colPath.slice(0, -1)], di));
      case 'index': {
        const [g, rt, ct] = [valueAt([], di), valueAt(rowPath, di), valueAt(colPath, di)];
        return typeof g === 'number' && typeof rt === 'number' && typeof ct === 'number' && rt * ct !== 0 ? (value * g) / (rt * ct) : { error: '#DIV/0!' };
      }
    }
    const bf = d.baseField ?? 0;
    const at = all.findIndex(([f]) => f === bf);
    if (at < 0) return null;
    const order = orderOf(bf);
    const mine = order.indexOf(all[at][1]);
    const swap = (k: string) => all.map(([f, x], j): [number, string] => [f, j === at ? k : x]);
    const sibling = (k: string) => valueAt(swap(k), di);
    switch (d.showAs) {
      case 'difference':
      case 'percent':
      case 'percentDiff': {
        const base = d.baseItem === PREVIOUS_ITEM ? order[mine - 1] : d.baseItem === NEXT_ITEM ? order[mine + 1] : keyOf(fields[bf]?.items[d.baseItem ?? 0]?.value);
        // The base item itself, or one before the first or after the last: itself, the whole of it, or nothing to measure.
        if (base === undefined || base === all[at][1]) return d.showAs === 'percent' ? 1 : null;
        const of = sibling(base);
        if (d.showAs === 'difference') return value - (typeof of === 'number' ? of : 0);
        if (typeof of !== 'number') return { error: '#N/A' };
        return d.showAs === 'percent' ? ratio(value, of) : ratio(value - of, of);
      }
      case 'runTotal':
      case 'percentOfRunningTotal': {
        const each = order.map((k) => sibling(k));
        const sum = (list: (EngineValue | undefined)[]) => list.reduce<number>((a, x) => a + (typeof x === 'number' ? x : 0), 0);
        const run = sum(each.slice(0, mine + 1));
        return d.showAs === 'runTotal' ? run : ratio(run, sum(each));
      }
      case 'percentOfParent': {
        const parent = at < rowPath.length ? [...rowPath.slice(0, at + 1), ...colPath] : [...rowPath, ...colPath.slice(0, at - rowPath.length + 1)];
        return ratio(value, valueAt(parent, di));
      }
      default: {
        // Ranks among the base field's items alike in everything else, ties sharing a rank.
        const others = order.map((k) => sibling(k)).filter((x): x is number => typeof x === 'number');
        return 1 + others.filter((x) => (d.showAs === 'rankAscending' ? x < value : x > value)).length;
      }
    }
  };

  // Each column's records as a set, to find quickly which of a row's fall in it.
  const colSets = colsOut.map((c) => (c.recs === keep ? undefined : new Set(c.recs)));
  rowsOut.forEach((r, ri) => {
    const row = firstDataRow + ri;
    for (const [col, text] of r.labels) put(row, col, text, r.kind === 'grand' ? 'grand' : r.kind === 'total' ? 'total' : 'label');
    if (r.blank) return;
    colsOut.forEach((c, ci) => {
      const d = data[r.data ?? c.data ?? 0];
      if (!d) return;
      const set = colSets[ci];
      const recs = set ? r.recs.filter((rec) => set.has(rec)) : r.recs;
      // Where nothing falls, Excel leaves the cell empty.
      if (!recs.length) return;
      const kind = r.kind === 'grand' || c.kind === 'grand' ? 'grand' : r.kind === 'total' || c.kind === 'total' ? 'total' : 'value';
      const value = summarize(d.fn, recs.map((rec) => rec[d.field]));
      const shown = d.showAs && d.showAs !== 'normal' ? showAs(value, r.path, c.path, r.data ?? c.data ?? 0) : value;
      if (shown !== null) put(row, labelCols + ci, shown, kind, shownFormat(d));
    });
  });

  // Report filters above the table, with a row between: each field's name and what it shows.
  model.pages.forEach((p, i) => {
    const row = -(model.pages.length + 1) + i;
    const pick = p.item !== undefined ? fields[p.field]?.items[p.item] : undefined;
    const several = !pick && fields[p.field]?.items.some((it) => it.hidden);
    put(row, 0, fields[p.field]?.name ?? '', 'page');
    put(row, 1, pick ? itemText(p.field, pick.value, pick.name) : several ? '(Multiple Items)' : '(All)', 'page');
  });

  return {
    height: firstDataRow + rowsOut.length,
    // With no values and nothing across, there is only the row labels' column.
    width: labelCols + (data.length || colLevels.length ? colsOut.length : 0),
    firstHeaderRow: 1,
    firstDataRow,
    firstDataCol: labelCols,
    cells,
  };
}

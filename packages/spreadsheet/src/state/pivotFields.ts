/**
 * What the PivotTable Fields pane does to a pivot table's fields, as Excel's
 * does it: ticking a field adds it to the values (numbers) or the rows,
 * dragging it moves it between the filters, columns, rows and values, and a
 * value's summary or a field's items can be changed. Each gives a new setup.
 */
import type { PivotFn, PivotSetup, PivotShowAs } from '../core/pivotTable';

export type PivotArea = 'pages' | 'cols' | 'rows' | 'data';

/** The Values field, which stands for the values among the rows or columns. */
export const VALUES_FIELD = -2;

export const PIVOT_FNS: PivotFn[] = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'];

/** A value's name for its summary of a field ("Sum of Units"), in the user's language. */
export type DataName = (fn: PivotFn, field: number) => string;

/** The fields in an area, in order: the values' by the field each summarises. */
export function areaFields(setup: PivotSetup, area: PivotArea): number[] {
  return area === 'data' ? setup.data.map((d) => d.field) : area === 'pages' ? setup.pages.map((p) => p.field) : setup[area];
}

/** The fields the pivot table uses, which the pane shows ticked. */
export function fieldsInUse(setup: PivotSetup): Set<number> {
  return new Set([...setup.rows, ...setup.cols, ...setup.pages.map((p) => p.field), ...setup.data.map((d) => d.field)].filter((f) => f >= 0));
}

/** A name not yet among the values', as Excel numbers a second "Sum of Units" "Sum of Units2". */
function unique(name: string, taken: string[]): string {
  if (!taken.includes(name)) return name;
  let n = 2;
  while (taken.includes(`${name}${n}`)) n++;
  return `${name}${n}`;
}

/** The Values field among the rows or columns only with several values, as Excel shows it. */
function tidy(setup: PivotSetup): PivotSetup {
  const several = setup.data.length > 1;
  const has = setup.rows.includes(VALUES_FIELD) || setup.cols.includes(VALUES_FIELD);
  if (several && !has) return { ...setup, cols: [...setup.cols, VALUES_FIELD] };
  if (!several && has) return { ...setup, rows: setup.rows.filter((f) => f !== VALUES_FIELD), cols: setup.cols.filter((f) => f !== VALUES_FIELD) };
  return setup;
}

function without(setup: PivotSetup, area: PivotArea, index: number): PivotSetup {
  if (area === 'data') return { ...setup, data: setup.data.filter((_, i) => i !== index) };
  if (area === 'pages') return { ...setup, pages: setup.pages.filter((_, i) => i !== index) };
  return { ...setup, [area]: setup[area].filter((_, i) => i !== index) };
}

function insert<T>(list: T[], at: number, item: T): T[] {
  const i = Math.max(0, Math.min(at, list.length));
  return [...list.slice(0, i), item, ...list.slice(i)];
}

/** A field ticked: numbers among the values, summed; anything else down the rows. */
export function addField(setup: PivotSetup, field: number, numeric: boolean, name: DataName): PivotSetup {
  return numeric ? moveTo(setup, field, 'data', setup.data.length, name, 'sum') : moveTo(setup, field, 'rows', setup.rows.length, name);
}

/** A field unticked: out of every area, its items' filter gone with it. */
export function removeField(setup: PivotSetup, field: number): PivotSetup {
  const hidden = { ...setup.hidden };
  delete hidden[field];
  return tidy({
    rows: setup.rows.filter((f) => f !== field),
    cols: setup.cols.filter((f) => f !== field),
    pages: setup.pages.filter((p) => p.field !== field),
    data: setup.data.filter((d) => d.field !== field),
    hidden,
  });
}

/**
 * A field put in an area at `at`. Among the filters, columns and rows a
 * field is in one only, so it leaves the others; among the values it may be
 * more than once, each summarised its own way (`fn`, or Sum).
 */
export function moveTo(setup: PivotSetup, field: number, area: PivotArea, at: number, name: DataName, fn: PivotFn = 'sum'): PivotSetup {
  if (area === 'data') {
    if (field === VALUES_FIELD) return setup;
    const entry = { field, fn, name: unique(name(fn, field), setup.data.map((d) => d.name)) };
    return tidy({ ...setup, data: insert(setup.data, at, entry) });
  }
  if (field === VALUES_FIELD && area === 'pages') return setup;
  const rest: PivotSetup = { ...setup, rows: setup.rows.filter((f) => f !== field), cols: setup.cols.filter((f) => f !== field), pages: setup.pages.filter((p) => p.field !== field) };
  // Where it was in the same area, the place it goes to counts without it.
  const before = areaFields(setup, area).indexOf(field);
  const to = before >= 0 && before < at ? at - 1 : at;
  if (area === 'pages') return tidy({ ...rest, pages: insert(rest.pages, to, { field }) });
  return tidy({ ...rest, [area]: insert(rest[area], to, field) });
}

/** The field in an area's `index`th place taken out of it, as its Remove Field does; out of the rows, columns or filters, its items' filter goes too. */
export function removeChip(setup: PivotSetup, area: PivotArea, index: number): PivotSetup {
  const field = areaFields(setup, area)[index];
  if (field === undefined || field === VALUES_FIELD) return setup;
  const next = without(setup, area, index);
  if (area !== 'data') {
    next.hidden = { ...next.hidden };
    delete next.hidden[field];
  }
  return tidy(next);
}

/** The field in an area's `index`th place moved to another area or place, as dragging it does. */
export function moveChip(setup: PivotSetup, from: { area: PivotArea; index: number }, to: { area: PivotArea; at: number }, name: DataName, fn: PivotFn = 'sum'): PivotSetup {
  const field = areaFields(setup, from.area)[from.index];
  if (field === undefined) return setup;
  if (from.area === 'data' && to.area === 'data') {
    const entry = setup.data[from.index];
    const rest = setup.data.filter((_, i) => i !== from.index);
    return { ...setup, data: insert(rest, to.at > from.index ? to.at - 1 : to.at, entry) };
  }
  // A value dragged out of the values leaves them; a field dragged into them leaves its area, as in Excel.
  if (from.area === 'data') return moveTo(tidy(without(setup, 'data', from.index)), field, to.area, to.at, name);
  if (to.area === 'data') return moveTo(tidy(without(setup, from.area, from.index)), field, 'data', to.at, name, fn);
  return moveTo(setup, field, to.area, to.at, name);
}

/** A value summarised another way; its name follows when it was the summary's own. */
export function setSummary(setup: PivotSetup, index: number, fn: PivotFn, name: DataName): PivotSetup {
  const d = setup.data[index];
  if (!d || d.fn === fn) return setup;
  const own = d.name.replace(/\d+$/, '') === name(d.fn, d.field);
  const taken = setup.data.filter((_, i) => i !== index).map((x) => x.name);
  return { ...setup, data: setup.data.map((x, i) => (i === index ? { ...x, fn, name: own ? unique(name(fn, d.field), taken) : x.name } : x)) };
}

/**
 * The items a field shows: the keys of those ticked, out of all of them.
 * A report filter showing one item picks it, as Excel's does; showing
 * several, it hides the rest.
 */
export function setShown(setup: PivotSetup, field: number, shown: string[], all: string[]): PivotSetup {
  const hidden = { ...setup.hidden };
  const keep = new Set(shown);
  const off = all.filter((k) => !keep.has(k));
  const page = setup.pages.findIndex((p) => p.field === field);
  if (page >= 0) {
    const one = shown.length === 1 && all.length > 1;
    delete hidden[field];
    if (!one && off.length) hidden[field] = off;
    return { ...setup, hidden, pages: setup.pages.map((p, i) => (i === page ? (one ? { field, item: shown[0] } : { field }) : p)) };
  }
  delete hidden[field];
  if (off.length) hidden[field] = off;
  return { ...setup, hidden };
}

/** The keys of the items a field shows now. */
export function shownItems(setup: PivotSetup, field: number, all: string[]): string[] {
  const pick = setup.pages.find((p) => p.field === field)?.item;
  if (pick !== undefined) return all.filter((k) => k === pick);
  const hidden = new Set(setup.hidden[field] ?? []);
  return all.filter((k) => !hidden.has(k));
}

/** A value shown another way, as Show Values As does: against `baseField` and `baseItem` (a key, or 'previous' or 'next') where the kind needs them. */
export function setShowAs(setup: PivotSetup, index: number, showAs: PivotShowAs, baseField?: number, baseItem?: string): PivotSetup {
  const d = setup.data[index];
  if (!d) return setup;
  const next = showAs === 'normal' ? { field: d.field, fn: d.fn, name: d.name } : { ...d, showAs, baseField: baseField ?? d.baseField ?? 0, baseItem: baseItem ?? d.baseItem ?? 'previous' };
  return { ...setup, data: setup.data.map((x, i) => (i === index ? next : x)) };
}

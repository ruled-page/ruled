import { msg, type Text } from '../i18n/message';
import type { OutlineAxis } from '../core/types';

export type Axis = 'rows' | 'cols';

/** A group: a run of rows (or columns) at its level or deeper. */
export interface OutlineGroup {
  level: number;
  start: number;
  end: number;
}

/** How deep Excel lets groups go. */
export const MAX_OUTLINE_LEVEL = 7;

/**
 * A change to one direction's outline, each part listing only what it
 * changes: levels, marks on collapsed groups' summaries, and the lines it
 * hides or shows.
 */
export interface OutlineChange {
  levels: Record<number, number>;
  collapsed: Record<number, boolean>;
  hidden: Record<number, boolean>;
}

const none = (): OutlineChange => ({ levels: {}, collapsed: {}, hidden: {} });
const cache = new WeakMap<Record<number, number>, OutlineGroup[]>();

/** The groups an outline's levels make, in order of where they start, each before the groups inside it. */
export function outlineGroups(o: OutlineAxis | undefined): OutlineGroup[] {
  if (!o) return [];
  let groups = cache.get(o.levels);
  if (groups) return groups;
  groups = [];
  const open: OutlineGroup[] = [];
  let prev = -2;
  const close = (depth: number) => {
    while (open.length > depth) open.pop()!.end = prev;
  };
  for (const line of Object.keys(o.levels).map(Number).sort((a, b) => a - b)) {
    const level = o.levels[line];
    if (line !== prev + 1) close(0);
    close(level);
    while (open.length < level) {
      const g = { level: open.length + 1, start: line, end: line };
      open.push(g);
      groups.push(g);
    }
    prev = line;
  }
  close(0);
  cache.set(o.levels, groups);
  return groups;
}

/** The deepest level of the outline's groups; 0 without any. */
export function outlineDepth(o: OutlineAxis | undefined): number {
  return outlineGroups(o).reduce((n, g) => Math.max(n, g.level), 0);
}

/** Width of each level's lane in the outline bars, in pixels. */
export const OUTLINE_LANE = 16;

/** How wide the bar beside the row numbers (or how tall the one above the column letters) is: a lane per level button; none without groups. */
export function outlineBarSize(o: OutlineAxis | undefined): number {
  const depth = outlineDepth(o);
  return depth ? (depth + 1) * OUTLINE_LANE + 4 : 0;
}

/** The row below a group (or above it), or the column right of it (or left), that sums it up and holds its button. */
export function summaryLine(o: OutlineAxis | undefined, g: OutlineGroup): number {
  return (o?.summaryAfter ?? true) ? g.end + 1 : g.start - 1;
}

/**
 * The lines one level deeper in groups (step 1), or one level out of them
 * (step -1), as Group and Ungroup do; or why Excel would refuse.
 */
export function regroup(o: OutlineAxis | undefined, axis: Axis, lines: number[], step: 1 | -1): OutlineChange | { refused: Text } {
  const change = none();
  for (const line of lines) {
    const next = (o?.levels[line] ?? 0) + step;
    if (next > MAX_OUTLINE_LEVEL) return { refused: msg('outline.refuse.tooDeep') };
    if (next >= 0) change.levels[line] = next;
  }
  if (step < 0 && !Object.keys(change.levels).length) return { refused: msg(axis === 'rows' ? 'outline.refuse.notGroupedRows' : 'outline.refuse.notGroupedColumns') };
  // A collapsed mark stays only on a line that still sums up a group.
  const levels = { ...o?.levels };
  for (const [line, v] of Object.entries(change.levels)) {
    if (v) levels[Number(line)] = v;
    else delete levels[Number(line)];
  }
  const after: OutlineAxis = { levels, collapsed: [], summaryAfter: o?.summaryAfter ?? true };
  const summaries = new Set(outlineGroups(after).map((g) => summaryLine(after, g)));
  for (const line of o?.collapsed ?? []) if (!summaries.has(line)) change.collapsed[line] = false;
  return change;
}

/** Hides a group's lines, marking its summary as collapsed. `hidden` is the lines hidden now. */
export function collapseGroup(o: OutlineAxis | undefined, hidden: ReadonlySet<number>, g: OutlineGroup): OutlineChange {
  const change = none();
  for (let line = g.start; line <= g.end; line++) if (!hidden.has(line)) change.hidden[line] = true;
  const s = summaryLine(o, g);
  if (s >= 0 && !o?.collapsed.includes(s)) change.collapsed[s] = true;
  return change;
}

/**
 * Shows a group's lines, except those in groups inside it that are
 * collapsed themselves, as Excel does. A group inside it summed up on the
 * same line opens with it, as that line's mark can't tell the two apart.
 */
export function expandGroup(o: OutlineAxis | undefined, hidden: ReadonlySet<number>, g: OutlineGroup): OutlineChange {
  const change = none();
  const s = summaryLine(o, g);
  const marked = new Set(o?.collapsed);
  if (marked.has(s)) change.collapsed[s] = false;
  const closed = outlineGroups(o).filter((x) => x.level > g.level && x.start >= g.start && x.end <= g.end && summaryLine(o, x) !== s && marked.has(summaryLine(o, x)));
  for (let line = g.start; line <= g.end; line++) {
    if (hidden.has(line) && !closed.some((x) => line >= x.start && line <= x.end)) change.hidden[line] = false;
  }
  return change;
}

/**
 * Shows the outline to level n, as Excel's level buttons do: lines in
 * groups that deep or deeper are hidden and those groups collapsed; the rest
 * of the grouped lines are shown.
 */
export function showLevel(o: OutlineAxis | undefined, hidden: ReadonlySet<number>, n: number): OutlineChange {
  const change = none();
  for (const [key, level] of Object.entries(o?.levels ?? {})) {
    const line = Number(key);
    if (level >= n && !hidden.has(line)) change.hidden[line] = true;
    else if (level < n && hidden.has(line)) change.hidden[line] = false;
  }
  const want = new Map<number, boolean>();
  for (const g of outlineGroups(o)) {
    const s = summaryLine(o, g);
    if (s >= 0) want.set(s, !!want.get(s) || g.level >= n);
  }
  const marked = new Set(o?.collapsed);
  for (const [line, c] of want) if (c !== marked.has(line)) change.collapsed[line] = c;
  return change;
}

/** Takes every line out of its groups. Lines the groups hid stay hidden, as in Excel. */
export function clearOutline(o: OutlineAxis | undefined): OutlineChange {
  const change = none();
  for (const key of Object.keys(o?.levels ?? {})) change.levels[Number(key)] = 0;
  for (const line of o?.collapsed ?? []) change.collapsed[line] = false;
  return change;
}

/**
 * The group Hide detail collapses at a line: the innermost one holding it,
 * else the deepest it sums up that is open. `shown(g)` says whether any of
 * a group's lines show.
 */
export function groupToHide(o: OutlineAxis | undefined, line: number, shown: (g: OutlineGroup) => boolean): OutlineGroup | undefined {
  const groups = outlineGroups(o);
  const inner = (list: OutlineGroup[]) => list.reduce<OutlineGroup | undefined>((a, g) => (!a || g.level > a.level ? g : a), undefined);
  return inner(groups.filter((g) => g.start <= line && line <= g.end)) ?? inner(groups.filter((g) => summaryLine(o, g) === line && shown(g)));
}

/** The group Show detail expands at a line: the outermost it sums up that is collapsed. */
export function groupToShow(o: OutlineAxis | undefined, line: number, shown: (g: OutlineGroup) => boolean): OutlineGroup | undefined {
  return outlineGroups(o).find((g) => summaryLine(o, g) === line && !shown(g));
}

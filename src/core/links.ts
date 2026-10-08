import { parseRange, rangeToA1 } from './refs';
import type { Hyperlink, Rect } from './types';

/**
 * Hyperlinks as the editor works with them: where a link goes, whether it
 * may be opened, and how giving cells a link (or taking theirs away) changes
 * a sheet's list of links. Nothing here touches the package, so the page
 * uses it as well as the worker.
 */

/** Where following a link goes. */
export type LinkDestination =
  /** A place in this workbook: "Sheet2!A1", a range, or a defined name. */
  | { kind: 'place'; location: string }
  /** A web page or an email, which may be opened; `url` is the address as the browser reads it. */
  | { kind: 'web' | 'email'; url: string }
  /** Anything else: a file, a script, another workbook or program. Never opened. */
  | { kind: 'refused'; address: string };

/**
 * Whether an address outside the workbook may be opened, and as what. Only
 * web pages (http and https) and email (mailto) may: the workbook is
 * untrusted, and javascript:, data:, file: and UNC paths can run code or
 * reach the user's own files. The address is read as the browser would read
 * it, which drops tabs and line breaks anywhere and spaces and control
 * characters at either end, so none of those can hide a scheme. `url` is the
 * address as the browser would open it, with an international domain in the
 * form that shows lookalike letters for what they are.
 */
export function checkAddress(address: string): LinkDestination {
  let url: URL;
  try {
    url = new URL(address);
  } catch {
    return { kind: 'refused', address };
  }
  if (url.protocol === 'mailto:') return { kind: 'email', url: url.href };
  if ((url.protocol === 'http:' || url.protocol === 'https:') && url.hostname) return { kind: 'web', url: url.href };
  return { kind: 'refused', address };
}

/**
 * A link's address and location as they are meant: an address starting "#"
 * is a place in the workbook (as in a HYPERLINK formula, and as some programs
 * write links to places), and a location may start with "#" too.
 */
function meant(link: { target?: string; location?: string }): { target?: string; location?: string } {
  const target = link.target?.trim();
  const inside = target?.startsWith('#');
  const location = (inside ? target : link.location?.trim())?.replace(/^#/, '');
  return { target: inside ? undefined : target || undefined, location: location || undefined };
}

/**
 * Where a link goes: a place in the workbook, or an address outside it (with
 * a location after an address naming a part of that page).
 */
export function linkDestination(link: { target?: string; location?: string }): LinkDestination | undefined {
  const { target, location } = meant(link);
  if (!target) return location ? { kind: 'place', location } : undefined;
  return checkAddress(location ? `${target}#${location}` : target);
}

/** A link's address as people read it: where it goes outside the workbook, or the place in it. */
export function linkAddress(link: { target?: string; location?: string }): string {
  const { target, location } = meant(link);
  if (!target) return location ?? '';
  return location ? `${target}#${location}` : target;
}

/** A sheet and the cells on it that a link's location names, before the sheet's name is looked up. */
export interface PlaceRef {
  /** The sheet as written, unquoted; unset for the sheet the link is on. */
  sheet?: string;
  /** The cells, or a defined name. */
  ref: string;
}

/**
 * Splits "Sheet2!A1", "'Q1 ''24'!B2:C4" or "A1" into the sheet and what
 * follows it, reading quotes as Excel writes them. Undefined when the text
 * isn't in that form.
 */
export function splitPlace(text: string): PlaceRef | undefined {
  const s = text.trim();
  if (s.startsWith("'")) {
    let name = '';
    let i = 1;
    for (; i < s.length; i++) {
      if (s[i] !== "'") name += s[i];
      else if (s[i + 1] === "'") name += s[i++];
      else break;
    }
    if (s[i] !== "'" || s[i + 1] !== '!' || !name) return undefined;
    const ref = s.slice(i + 2).trim();
    return ref && !ref.includes('!') ? { sheet: name, ref } : undefined;
  }
  const bang = s.indexOf('!');
  if (bang < 0) return s ? { ref: s } : undefined;
  const sheet = s.slice(0, bang).trim();
  const ref = s.slice(bang + 1).trim();
  return sheet && ref && !ref.includes('!') ? { sheet, ref } : undefined;
}

/** A defined name, as the workbook keeps it. */
export interface PlaceName {
  name: string;
  formula: string;
  localSheet?: number;
}

/**
 * The sheet and cells a link's location names: a range on a sheet, a range
 * on the sheet the link is on, or a defined name that stands for a range
 * (a name local to that sheet first). Undefined when the place isn't in the
 * workbook, or is a name that is worked out by a formula.
 */
export function resolvePlace(location: string, sheets: readonly string[], names: readonly PlaceName[], own: number): { sheet: number; rect: Rect } | undefined {
  const sheetIndex = (name: string) => sheets.findIndex((s) => s.toLowerCase() === name.toLowerCase());
  const at = (text: string, from: number, depth: number): { sheet: number; rect: Rect } | undefined => {
    const place = splitPlace(text);
    if (!place) return undefined;
    const sheet = place.sheet === undefined ? from : sheetIndex(place.sheet);
    if (sheet < 0) return undefined;
    const rect = parseRange(place.ref);
    if (rect) return { sheet, rect };
    if (place.sheet !== undefined || depth > 0) return undefined;
    const lower = place.ref.toLowerCase();
    const name = names.find((n) => n.name.toLowerCase() === lower && n.localSheet === from) ?? names.find((n) => n.name.toLowerCase() === lower && n.localSheet === undefined);
    return name ? at(name.formula, name.localSheet ?? from, depth + 1) : undefined;
  };
  return at(location.replace(/^=/, ''), own, 0);
}

/**
 * Whether a formula (in file form) is a HYPERLINK call, which makes its cell
 * a link to the address the call works out.
 */
export function isHyperlinkFormula(formula: string | undefined): boolean {
  if (!formula) return false;
  let i = 0;
  while (formula.charCodeAt(i) === 32) i++;
  if (formula.slice(i, i + 9).toUpperCase() !== 'HYPERLINK') return false;
  i += 9;
  while (formula.charCodeAt(i) === 32) i++;
  return formula[i] === '(';
}

/**
 * A HYPERLINK formula's address when it is written into the formula as text,
 * as in =HYPERLINK("https://example.com", "Example"); undefined when it is
 * worked out from cells, which only the calculation engine can do.
 */
export function literalHyperlinkTarget(formula: string): string | undefined {
  if (!isHyperlinkFormula(formula)) return undefined;
  let i = formula.indexOf('(') + 1;
  while (formula.charCodeAt(i) === 32) i++;
  if (formula[i] !== '"') return undefined;
  let text = '';
  for (i++; i < formula.length; i++) {
    if (formula[i] !== '"') text += formula[i];
    else if (formula[i + 1] === '"') text += formula[i++];
    else break;
  }
  if (formula[i] !== '"') return undefined;
  i++;
  while (formula.charCodeAt(i) === 32) i++;
  return formula[i] === ',' || formula[i] === ')' ? text : undefined;
}

const sameRect = (a: Rect, b: Rect) => a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
const overlaps = (a: Rect, b: Rect) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;

/** Two links that are the same, from the same element of the file if either is. */
export function sameLink(a: Hyperlink, b: Hyperlink): boolean {
  return sameRect(a.rect, b.rect) && a.target === b.target && a.location === b.location && a.display === b.display && a.tooltip === b.tooltip && a.id === b.id;
}

/** The parts of `a` outside `b`: bands above and below, then pieces either side. */
function subtract(a: Rect, b: Rect): Rect[] {
  if (!overlaps(a, b)) return [a];
  const out: Rect[] = [];
  if (a.r1 < b.r1) out.push({ ...a, r2: b.r1 - 1 });
  if (a.r2 > b.r2) out.push({ ...a, r1: b.r2 + 1 });
  const r1 = Math.max(a.r1, b.r1);
  const r2 = Math.min(a.r2, b.r2);
  if (a.c1 < b.c1) out.push({ r1, r2, c1: a.c1, c2: b.c1 - 1 });
  if (a.c2 > b.c2) out.push({ r1, r2, c1: b.c2 + 1, c2: a.c2 });
  return out;
}

/** The fields of a link that say where it goes and how it reads, without its cells. */
export type LinkFields = Pick<Hyperlink, 'target' | 'location' | 'display' | 'tooltip'>;

const fieldsOf = (l: Hyperlink): LinkFields => {
  const out: LinkFields = {};
  if (l.target !== undefined) out.target = l.target;
  if (l.location !== undefined) out.location = l.location;
  if (l.display !== undefined) out.display = l.display;
  if (l.tooltip !== undefined) out.tooltip = l.tooltip;
  return out;
};

/**
 * The change that gives every cell in `rects` the link `link`, or takes their
 * links away when there is none: links that cover any of those cells are
 * removed, and the parts of them outside the cells added back.
 */
export function replaceLinks(links: readonly Hyperlink[], rects: readonly Rect[], link?: LinkFields): { add: Hyperlink[]; remove: Hyperlink[] } {
  const add: Hyperlink[] = [];
  const remove: Hyperlink[] = [];
  for (const l of links) {
    if (!rects.some((r) => overlaps(l.rect, r))) continue;
    remove.push(l);
    let parts = [l.rect];
    for (const r of rects) parts = parts.flatMap((p) => subtract(p, r));
    for (const rect of parts) add.push({ rect, ...fieldsOf(l) });
  }
  if (link) for (const rect of rects) add.push({ rect, ...link });
  return { add, remove };
}

/** At most this many links are made by one copy; a few cells copied over a whole column carry their links that far, not one per copy. */
const MAX_COPIED = 10_000;

/**
 * The change that copies the links in `src` (from `from`, the source sheet's
 * links) to `dst`, tiling it as cells are copied: the links `dst` had go,
 * as Excel's paste replaces them. A link that covers all of `src` covers all
 * of `dst`.
 */
export function copyLinks(from: readonly Hyperlink[], src: Rect, to: readonly Hyperlink[], dst: Rect): { add: Hyperlink[]; remove: Hyperlink[] } {
  const change = replaceLinks(to, [dst]);
  const h = src.r2 - src.r1 + 1;
  const w = src.c2 - src.c1 + 1;
  let made = 0;
  for (const l of from) {
    if (!overlaps(l.rect, src)) continue;
    const part = { r1: Math.max(l.rect.r1, src.r1), c1: Math.max(l.rect.c1, src.c1), r2: Math.min(l.rect.r2, src.r2), c2: Math.min(l.rect.c2, src.c2) };
    if (sameRect(part, src)) {
      change.add.push({ rect: { ...dst }, ...fieldsOf(l) });
      continue;
    }
    for (let top = dst.r1; top <= dst.r2 && made < MAX_COPIED; top += h) {
      for (let left = dst.c1; left <= dst.c2 && made < MAX_COPIED; left += w) {
        const rect = {
          r1: top + part.r1 - src.r1,
          c1: left + part.c1 - src.c1,
          r2: Math.min(dst.r2, top + part.r2 - src.r1),
          c2: Math.min(dst.c2, left + part.c2 - src.c1),
        };
        if (rect.r1 > dst.r2 || rect.c1 > dst.c2) continue;
        change.add.push({ rect, ...fieldsOf(l) });
        made++;
      }
    }
  }
  return change;
}

/**
 * The change that moves links with their rows when `rect`'s rows are sorted:
 * `rowFrom[i]` is the row whose cells now sit in row `top + i`. Links on one
 * row inside the range move with it; links across rows, or reaching outside
 * the range, stay where they are.
 */
export function sortLinks(links: readonly Hyperlink[], rect: Rect, top: number, rowFrom: readonly number[]): { add: Hyperlink[]; remove: Hyperlink[] } {
  const to = new Map<number, number>();
  rowFrom.forEach((from, i) => {
    if (from !== top + i) to.set(from, top + i);
  });
  const add: Hyperlink[] = [];
  const remove: Hyperlink[] = [];
  for (const l of links) {
    const row = to.get(l.rect.r1);
    if (row === undefined || l.rect.r2 !== l.rect.r1 || l.rect.c1 < rect.c1 || l.rect.c2 > rect.c2) continue;
    remove.push(l);
    add.push({ ...fieldsOf(l), rect: { ...l.rect, r1: row, r2: row } });
  }
  return { add, remove };
}

/** A sheet's links after a change: the ones removed go, the ones added that it doesn't have already come after the rest. */
export function applyLinkChange(links: readonly Hyperlink[], change: { add: readonly Hyperlink[]; remove: readonly Hyperlink[] }): { links: Hyperlink[]; added: Hyperlink[]; removed: Hyperlink[] } {
  const removed = links.filter((l) => change.remove.some((r) => sameLink(r, l)));
  const kept = links.filter((l) => !removed.includes(l));
  const added = change.add.filter((a, i) => !kept.some((l) => sameLink(l, a)) && change.add.findIndex((b) => sameLink(a, b)) === i);
  return { links: [...kept, ...added], added, removed };
}

/** How a sheet's links differ from the ones it had: by range, the address each now goes to, or null where its link went. */
export function linkDifferences(before: readonly Hyperlink[], after: readonly Hyperlink[]): { ref: string; address: string | null; was: boolean }[] {
  const removed = before.filter((b) => !after.some((a) => sameLink(a, b)));
  const added = after.filter((a) => !before.some((b) => sameLink(a, b)));
  const out = new Map<string, { ref: string; address: string | null; was: boolean }>();
  for (const l of removed) out.set(rangeToA1(l.rect), { ref: rangeToA1(l.rect), address: null, was: true });
  for (const l of added) {
    const ref = rangeToA1(l.rect);
    out.set(ref, { ref, address: linkAddress(l), was: out.has(ref) });
  }
  return [...out.values()];
}

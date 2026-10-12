import { parseA1Ref, stringifyTokens, tokenize } from '@borgar/fx';
import { sameName } from './names';
import type { DefinedName, NameRef } from './types';
import { elements, escapeXml, firstElement, parseAttrs } from './xml';

/** Whether two names are defined alike: name, scope, formula, comment and visibility. */
export function sameDefinition(a: DefinedName, b: DefinedName): boolean {
  return a.name === b.name && a.formula === b.formula && a.localSheet === b.localSheet && (a.comment ?? '') === (b.comment ?? '') && !!a.hidden === !!b.hidden;
}

/**
 * A formula with the name `from` renamed `to` wherever it means that name,
 * or undefined when it doesn't use it. On sheet `own` (undefined for a name
 * for the whole workbook), "Name" means one local to that sheet before the
 * workbook's, and "Sheet!Name" one local to that sheet before the
 * workbook's. `names` are the names before the rename.
 */
export function renameInFormula(formula: string, from: NameRef, to: string, own: number | undefined, names: DefinedName[], sheetNames: string[]): string | undefined {
  if (!formula.toUpperCase().includes(from.name.toUpperCase())) return undefined;
  let list;
  try {
    list = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return undefined;
  }
  let renamed = false;
  const out = list.map((t) => {
    if (t.type !== 'range_named') return t;
    const ref = parseA1Ref(t.value, { allowTernary: false });
    if (!ref || !('name' in ref) || !ref.name || !sameName(ref.name, from.name) || (ref.context?.length ?? 0) > 1) return t;
    const context = ref.context?.[0];
    const sheet = context === undefined ? own : sheetNames.findIndex((s) => sameName(s, context));
    if (sheet === -1) return t;
    const local = sheet !== undefined && names.some((n) => n.localSheet === sheet && sameName(n.name, from.name));
    if ((local ? sheet : undefined) !== from.localSheet) return t;
    renamed = true;
    return { ...t, value: t.value.slice(0, t.value.length - ref.name.length) + to };
  });
  return renamed ? stringifyTokens(out) : undefined;
}

/** An attribute value escaped so that line breaks and tabs, as comments may hold, survive being read back. */
function escapeAttr(s: string): string {
  return escapeXml(s).replace(/[\n\r\t]/g, (c) => `&#${c.charCodeAt(0)};`);
}

/** A definedName element for a name, keeping attributes it had that the editor doesn't change. */
function definedNameXml(prefix: string, n: DefinedName, kept: Record<string, string>): string {
  const attrs: Record<string, string> = { ...kept, name: n.name };
  if (n.comment) attrs.comment = n.comment;
  else delete attrs.comment;
  if (n.localSheet !== undefined) attrs.localSheetId = String(n.localSheet);
  else delete attrs.localSheetId;
  if (n.hidden) attrs.hidden ??= '1';
  const text = Object.entries(attrs)
    .map(([k, v]) => ` ${k}="${escapeAttr(v)}"`)
    .join('');
  return `<${prefix}definedName${text}>${escapeXml(n.formula)}</${prefix}definedName>`;
}

/**
 * workbook.xml with its defined names as they now are. A name read from the
 * file keeps its element, byte for byte while it is unchanged and with its
 * other attributes when it changes; new names go at the end. The
 * definedNames element is added where the schema puts it when the file has
 * none, and removed when no names are left.
 */
export function writeDefinedNames(xml: string, file: DefinedName[], now: DefinedName[]): string {
  const block = firstElement(xml, 'definedNames');
  const qname = block?.qname ?? firstElement(xml, 'sheets')?.qname ?? '';
  const prefix = qname.slice(0, qname.lastIndexOf(':') + 1);
  const found = block?.inner ? elements(xml, 'definedName', block.start, block.end) : [];
  // Elements are matched to the names read from them by position; if that can't be done, every name is written afresh.
  const matched = found.length === file.length;
  const byIndex = new Map(now.filter((n) => n.fileIndex !== undefined).map((n) => [n.fileIndex!, n]));
  const parts: string[] = [];
  if (matched) {
    found.forEach((el, i) => {
      const n = byIndex.get(i);
      if (n) parts.push(sameDefinition(n, file[i]) ? xml.slice(el.start, el.end) : definedNameXml(prefix, n, parseAttrs(el.attrs)));
    });
  }
  for (const n of now) if (!matched || n.fileIndex === undefined) parts.push(definedNameXml(prefix, n, {}));

  if (block) {
    if (!parts.length) return xml.slice(0, block.start) + xml.slice(block.end);
    return xml.slice(0, block.start) + `<${block.qname}${block.attrs}>${parts.join('')}</${block.qname}>` + xml.slice(block.end);
  }
  if (!parts.length) return xml;
  // CT_Workbook puts definedNames straight after the last of these the workbook has.
  for (const before of ['externalReferences', 'functionGroups', 'sheets']) {
    const el = firstElement(xml, before);
    if (el) return xml.slice(0, el.end) + `<${prefix}definedNames>${parts.join('')}</${prefix}definedNames>` + xml.slice(el.end);
  }
  return xml;
}

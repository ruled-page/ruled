/**
 * Pivot tables whose fields were changed, written back: the cache made again
 * from the source's rows as they now are, and the definition with its fields
 * in their new places. Excel is asked to refresh the cache on opening, so it
 * works out the rows and columns of items itself.
 */
import { addRel, setOverride } from './commentWriter';
import { keyOf, sharedValue, type PivotModel, type PivotValue } from './pivotTable';
import { attr, child, children, elements, escapeXml, findTagEnd, firstElement, parseXml, replaceElements, setAttributes } from './xml';
import type { OoxmlPackage } from './zip';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

/** A date's serial number as the date-time a cache gives it. */
function isoOf(serial: number): string {
  return new Date(Math.round((serial - 25_569) * 86_400_000)).toISOString().slice(0, 19);
}

/** A value as a cache writes it, among a field's shared items or in a record. */
function valueXml(v: PivotValue, dates: boolean): string {
  if (v === null || v === '') return '<m/>';
  if (typeof v === 'number') return dates ? `<d v="${isoOf(v)}"/>` : `<n v="${v}"/>`;
  if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"/>`;
  if (typeof v === 'object') return `<e v="${escapeXml(v.error)}"/>`;
  return `<s v="${escapeXml(v)}"/>`;
}

/** What kinds of value a field's shared items hold, as Excel says it on `<sharedItems>`. */
function sharedAttrs(values: PivotValue[], dates: boolean, count: number | undefined): Record<string, string | null> {
  let str = false;
  let blank = false;
  let num = false;
  let int = true;
  let min = Infinity;
  let max = -Infinity;
  for (const v of values) {
    if (v === null || v === '') blank = true;
    else if (typeof v === 'number') {
      num = true;
      if (!Number.isInteger(v)) int = false;
      min = Math.min(min, v);
      max = Math.max(max, v);
    } else str = true;
  }
  const asDate = dates && num;
  return {
    containsSemiMixedTypes: str || blank ? null : '0',
    containsNonDate: asDate && !str ? '0' : null,
    containsDate: asDate ? '1' : null,
    containsString: str ? null : '0',
    containsBlank: blank ? '1' : null,
    containsMixedTypes: str && num ? '1' : null,
    containsNumber: num && !asDate ? '1' : null,
    containsInteger: num && !asDate && int ? '1' : null,
    minValue: num && !asDate ? String(min) : null,
    maxValue: num && !asDate ? String(max) : null,
    minDate: asDate ? isoOf(min) : null,
    maxDate: asDate ? isoOf(max) : null,
    longText: null,
    count: count === undefined ? null : String(count),
  };
}

/** The root element's start tag with attributes set. */
function setRoot(xml: string, name: string, attrs: Record<string, string | null>): string {
  const at = xml.search(new RegExp(`<(\\w+:)?${name}\\b`));
  if (at < 0) return xml;
  const end = findTagEnd(xml, at) + 1;
  return xml.slice(0, at) + setAttributes(xml.slice(at, end), attrs) + xml.slice(end);
}

export interface PivotCacheParts {
  definition: string;
  /** The records part, when the cache has one. */
  records?: string;
  /** A value's index among a field's shared items, or -1. */
  index: (field: number, value: PivotValue | undefined) => number;
  /** How many shared items a field has. */
  count: (field: number) => number;
}

/**
 * A pivot cache made again from `records`, the source's rows: each field's
 * shared items keep their order, with values new since after them, so the
 * items of other pivot tables sharing it still point at theirs. A field
 * lists its values when it did, when it holds anything but numbers, or when
 * it is in `listed` (on a pivot table's rows, columns or filters).
 */
export function rewritePivotCache(xml: string, hasRecords: boolean, records: PivotValue[][], listed: ReadonlySet<number>): PivotCacheParts {
  const doc = parseXml(xml);
  const fields = children(child(doc, 'cacheFields'), 'cacheField').map((cf, i) => {
    const shared = child(cf, 'sharedItems');
    const values = (shared?.children ?? []).map(sharedValue);
    const keys = new Map<string, number>();
    values.forEach((v, x) => {
      if (!keys.has(keyOf(v))) keys.set(keyOf(v), x);
    });
    const column = records.map((r) => r[i] ?? null);
    return {
      values,
      keys,
      column,
      dates: attr(shared, 'containsDate') === '1' || !!shared?.children.some((c) => c.name === 'd'),
      list: values.length > 0 || listed.has(i) || column.some((v) => v !== null && v !== '' && typeof v !== 'number'),
    };
  });
  for (const f of fields) {
    if (!f.list) continue;
    for (const v of f.column) {
      if (f.keys.has(keyOf(v))) continue;
      f.keys.set(keyOf(v), f.values.length);
      f.values.push(v);
    }
  }
  let i = 0;
  let definition = replaceElements(xml, 'sharedItems', (el) => {
    const f = fields[i++];
    if (!f) return `<${el.qname}${el.attrs}${el.inner === null ? '/>' : `>${el.inner}</${el.qname}>`}`;
    const tag = setAttributes(`<${el.qname}${el.attrs}>`, sharedAttrs(f.list ? f.values : f.column, f.dates, f.list ? f.values.length : undefined));
    return f.list ? `${tag}${f.values.map((v) => valueXml(v, f.dates)).join('')}</${el.qname}>` : `${tag.slice(0, -1)}/>`;
  });
  definition = setRoot(definition, 'pivotCacheDefinition', { refreshOnLoad: '1', ...(hasRecords && { recordCount: String(records.length) }) });
  const rows = hasRecords
    ? records.map((r) => `<r>${fields.map((f, k) => (f.list ? `<x v="${f.keys.get(keyOf(r[k] ?? null))}"/>` : valueXml(r[k] ?? null, f.dates))).join('')}</r>`)
    : [];
  return {
    definition,
    ...(hasRecords && {
      records: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="${MAIN_NS}" xmlns:r="${R_NS}" count="${records.length}">${rows.join('')}</pivotCacheRecords>`,
    }),
    index: (field, value) => fields[field]?.keys.get(keyOf(value ?? null)) ?? -1,
    count: (field) => fields[field]?.values.length ?? 0,
  };
}

const AXES = { row: 'axisRow', col: 'axisCol', page: 'axisPage' } as const;
/** Elements laid out from the fields, which follow `<pivotFields>`: written again, or left out for Excel to work out. */
const LAID_OUT = ['rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'formats', 'conditionalFormats', 'chartFormats', 'filters'];

/**
 * A pivot table's definition with its fields placed as `model` says: each
 * field's area, items and hidden items, the rows', columns', filters' and
 * values' lists, and where it now lies. Its formats, which point at items
 * and areas that may have moved, are left out, as are the rows and columns
 * of items, which Excel works out when it refreshes the cache.
 */
export function rewritePivotTable(
  xml: string,
  model: PivotModel,
  location: { ref: string; firstHeaderRow: number; firstDataRow: number; firstDataCol: number },
  cache: Pick<PivotCacheParts, 'index' | 'count'>,
): string {
  const dataFields = new Set(model.data.map((d) => d.field));
  // Where each field's items are among those written, for the filters' picks.
  const written: Map<number, number>[] = [];
  let i = 0;
  let out = replaceElements(xml, 'pivotField', (el) => {
    const n = i++;
    const field = model.fields[n];
    const pos = new Map<number, number>();
    written[n] = pos;
    if (!field) return `<${el.qname}${el.attrs}${el.inner === null ? '/>' : `>${el.inner}</${el.qname}>`}`;
    const tag = setAttributes(`<${el.qname}${el.attrs}>`, { axis: field.axis ? AXES[field.axis] : null, dataField: dataFields.has(n) ? '1' : null });
    let items = '';
    if (field.axis) {
      const list: string[] = [];
      const used = new Set<number>();
      field.items.forEach((it, k) => {
        const x = it.x ?? cache.index(n, it.value);
        if (x < 0 || used.has(x)) return;
        used.add(x);
        pos.set(k, list.length);
        list.push(`<item x="${x}"${it.hidden ? ' h="1"' : ''}${it.collapsed ? ' sd="0"' : ''}${it.name !== undefined ? ` n="${escapeXml(it.name)}"` : ''}/>`);
      });
      // Values the source no longer has, as Excel keeps listing them.
      for (let x = 0; x < cache.count(n); x++) if (!used.has(x)) list.push(`<item m="1" x="${x}"/>`);
      if (field.subtotals) list.push('<item t="default"/>');
      items = `<items count="${list.length}">${list.join('')}</items>`;
    }
    const rest = el.inner === null ? '' : replaceElements(el.inner, 'items', () => '');
    return items || rest ? `${tag}${items}${rest}</${el.qname}>` : `${tag.slice(0, -1)}/>`;
  });
  for (const name of LAID_OUT) out = replaceElements(out, name, () => '');
  const list = (name: string, xs: number[]) => (xs.length ? `<${name} count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join('')}</${name}>` : '');
  const pages = model.pages.length
    ? `<pageFields count="${model.pages.length}">${model.pages
        .map((p) => {
          const at = p.item !== undefined ? written[p.field]?.get(p.item) : undefined;
          return `<pageField fld="${p.field}"${at !== undefined ? ` item="${at}"` : ''} hier="-1"/>`;
        })
        .join('')}</pageFields>`
    : '';
  const data = model.data.length
    ? `<dataFields count="${model.data.length}">${model.data
        .map((d) => `<dataField name="${escapeXml(d.name)}" fld="${d.field}"${d.fn !== 'sum' ? ` subtotal="${d.fn}"` : ''} baseField="0" baseItem="0"${d.numFmtId !== undefined ? ` numFmtId="${d.numFmtId}"` : ''}/>`)
        .join('')}</dataFields>`
    : '';
  const fieldsEnd = elements(out, 'pivotFields', 0, out.length, 1)[0]?.end;
  if (fieldsEnd !== undefined) out = out.slice(0, fieldsEnd) + list('rowFields', model.rows) + list('colFields', model.cols) + pages + data + out.slice(fieldsEnd);
  out = replaceElements(out, 'location', (el) =>
    setAttributes(`<${el.qname}${el.attrs}/>`, {
      ref: location.ref,
      firstHeaderRow: String(location.firstHeaderRow),
      firstDataRow: String(location.firstDataRow),
      firstDataCol: String(location.firstDataCol),
      rowPageCount: model.pages.length ? String(model.pages.length) : null,
      colPageCount: model.pages.length ? '1' : null,
    }),
  );
  return setRoot(out, 'pivotTableDefinition', { dataOnRows: model.rows.includes(-2) ? '1' : null, dataPosition: null });
}

/** A new pivot table's cache before its values are in: its source and a field for each column. */
export function newPivotCacheXml(model: PivotModel): string {
  const src = model.source ?? {};
  const where = src.name !== undefined ? `name="${escapeXml(src.name)}"` : `ref="${escapeXml(src.ref ?? '')}" sheet="${escapeXml(src.sheet ?? '')}"`;
  const fields = model.fields.map((f) => `<cacheField name="${escapeXml(f.name)}" numFmtId="${f.numFmtId ?? 0}"><sharedItems/></cacheField>`).join('');
  return (
    `${XML_HEAD}<pivotCacheDefinition xmlns="${MAIN_NS}" xmlns:r="${R_NS}" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="0">` +
    `<cacheSource type="worksheet"><worksheetSource ${where}/></cacheSource><cacheFields count="${model.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`
  );
}

/** A new pivot table's definition before its fields are placed, with Excel's default style. */
export function newPivotTableXml(model: PivotModel, cacheId: number): string {
  return (
    `${XML_HEAD}<pivotTableDefinition xmlns="${MAIN_NS}" name="${escapeXml(model.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="0"/><pivotFields count="${model.fields.length}">${'<pivotField showAll="0"/>'.repeat(model.fields.length)}</pivotFields>` +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>'
  );
}

export interface NewPivotParts {
  sheet: number;
  cacheId: number;
  table: { path: string; xml: string };
  cache: { path: string; xml: string };
  records: { path: string; xml: string };
}

/** Workbook elements a `<pivotCaches>` goes before. */
const AFTER_PIVOT_CACHES = ['smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

/** The workbook's list of pivot caches with one more. */
function withPivotCache(xml: string, cacheId: number, id: string): string {
  const prefix = /<(\w+:)?workbook\b/.exec(xml)?.[1] ?? '';
  const ns = /<(\w+:)?workbook\b[^>]*\sxmlns:r=/.test(xml) ? '' : ` xmlns:r="${R_NS}"`;
  const tag = `<${prefix}pivotCache cacheId="${cacheId}" r:id="${id}"${ns}/>`;
  const caches = firstElement(xml, 'pivotCaches');
  if (caches) {
    const inner = `${caches.inner ?? ''}${tag}`;
    return `${xml.slice(0, caches.start)}<${caches.qname}${caches.attrs}>${inner}</${caches.qname}>${xml.slice(caches.end)}`;
  }
  const before = AFTER_PIVOT_CACHES.map((n) => firstElement(xml, n)?.start).filter((n): n is number => n !== undefined);
  const at = before.length ? Math.min(...before) : xml.lastIndexOf(`</${prefix}workbook>`);
  return `${xml.slice(0, at)}<${prefix}pivotCaches>${tag}</${prefix}pivotCaches>${xml.slice(at)}`;
}

/** A pivot table inserted here added to the package: its parts, their content types, and the relationships that find them. */
export function addPivotTable(pkg: OoxmlPackage, workbookPath: string, sheetPath: string, parts: NewPivotParts): void {
  pkg.setText(parts.records.path, parts.records.xml);
  setOverride(pkg, parts.records.path, `${CT}.pivotCacheRecords+xml`);
  const records = addRel(pkg, parts.cache.path, `${R_NS}/pivotCacheRecords`, parts.records.path);
  pkg.setText(parts.cache.path, setRoot(parts.cache.xml, 'pivotCacheDefinition', { 'r:id': records }));
  setOverride(pkg, parts.cache.path, `${CT}.pivotCacheDefinition+xml`);
  pkg.setText(parts.table.path, parts.table.xml);
  setOverride(pkg, parts.table.path, `${CT}.pivotTable+xml`);
  addRel(pkg, parts.table.path, `${R_NS}/pivotCacheDefinition`, parts.cache.path);
  addRel(pkg, sheetPath, `${R_NS}/pivotTable`, parts.table.path);
  const id = addRel(pkg, workbookPath, `${R_NS}/pivotCacheDefinition`, parts.cache.path);
  const workbook = pkg.text(workbookPath);
  if (workbook) pkg.setText(workbookPath, withPivotCache(workbook, parts.cacheId, id));
}

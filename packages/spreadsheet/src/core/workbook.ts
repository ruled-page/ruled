import { msg, TextError, type Message, type MessageKey, type Text } from '../i18n/message';
import { parseChart, type ChartContext, type ChartSpec } from './charts';
import { parseRange, rangeToA1, toA1 } from '@sitepen/formula-engine/refs';
import { parseSharedStrings, scanSheet, type CellBudget, type ScanResult } from './sheetScanner';
import { parseConditionalFormats } from './conditionalFormat';
import { readCustomProperties } from './properties';
import { StyleBook } from './styles';
import { builtinTableLook } from './tableStyles';
import { mergeNotes, parseNotes, parsePeople, parseThreads, type Person } from './comments';
import { readHyperlinks } from './hyperlinks';
import type { Anchor, BlockedRegion, CellNote, DefinedName, FilterCondition, FilterOperator, FilterState, Finding, Rect, SheetInfo, SheetObject } from './types';
import { parseValidations } from './validations';
import { attr, child, children, descendants, parseXml, textOf, type XmlNode } from './xml';
import type { OoxmlPackage, Relationship } from './zip';
import type { Iteration } from '@sitepen/formula-engine';

export type { DefinedName } from './types';

export interface SheetModel {
  info: SheetInfo;
  path: string;
  /** The worksheet's XML, decoded from the package each time it is read. */
  readonly xml?: string;
  scan?: ScanResult;
  /** Column ranges with a default style, applied to newly created cells. */
  colStyles: { min: number; max: number; style: number }[];
}

export interface WorkbookModel {
  workbookPath: string;
  sheets: SheetModel[];
  styles: StyleBook;
  names: DefinedName[];
  date1904: boolean;
  /** Iterative calculation, when the workbook turns it on: circular references are worked round rather than read as 0. */
  iteration?: Iteration;
  findings: Finding[];
  /** The workbook carries a VBA project. */
  /** A VBA project or Excel 4.0 macro sheets. */
  hasMacros: boolean;
  /** ActiveX controls or embedded (OLE) objects. */
  hasActiveContent: boolean;
  /** Custom document properties. */
  properties: Record<string, string>;
  /** The people who have written threaded comments. */
  people: Person[];
  /** The workbook's structure is protected: sheets can't be added, deleted, renamed, moved, hidden or unhidden. */
  structureLocked: boolean;
  /** What chart references are resolved against: the theme, the sheets and the names. */
  chartContext: ChartContext;
}

const EMU_PER_PX = 9525;
/** Hidden-column and width ranges are only expanded this far past the used range. */
const COLUMN_PADDING = 64;

const CHART_TYPES: Record<string, MessageKey> = {
  barChart: 'core.chartType.bar',
  bar3DChart: 'core.chartType.bar3D',
  lineChart: 'core.chartType.line',
  line3DChart: 'core.chartType.line3D',
  pieChart: 'core.chartType.pie',
  pie3DChart: 'core.chartType.pie3D',
  doughnutChart: 'core.chartType.doughnut',
  areaChart: 'core.chartType.area',
  area3DChart: 'core.chartType.area3D',
  scatterChart: 'core.chartType.scatter',
  bubbleChart: 'core.chartType.bubble',
  radarChart: 'core.chartType.radar',
  stockChart: 'core.chartType.stock',
  surfaceChart: 'core.chartType.surface',
  surface3DChart: 'core.chartType.surface3D',
  ofPieChart: 'core.chartType.pieOfPie',
};

/**
 * Texts joined into one list, "a, b, c" by default or "a; b; c" with
 * 'core.list.clauses': each item is translated before it is joined.
 */
export function joinTexts(items: Text[], key: MessageKey = 'core.list.items'): Text {
  return items.slice(0, -1).reduceRight<Text>((rest, first) => msg(key, { first, rest }), items[items.length - 1] ?? '');
}

/**
 * A filter's picked values by sheet column, as the cells show them ('' for
 * blanks); undefined when a column filters some other way (by condition,
 * top ten, colour or date groups), which the editor leaves as it is.
 */
/** The operators Excel writes on a custom filter; "equal" is left out. */
const FILTER_OPERATORS = new Set<string>(['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual']);

/**
 * A filter's columns: those filtered by values, and those by a condition the
 * editor can change (custom rules, Top 10, above or below the average).
 * Undefined when a column is filtered some other way (dates in a period, by
 * colour or icon), so that the filter is kept as the file has it.
 */
function filterColumns(node: XmlNode | undefined, rect: Rect): Pick<FilterState, 'columns' | 'conditions'> | undefined {
  const columns: Record<number, string[]> = {};
  const conditions: Record<number, FilterCondition> = {};
  const on = (v: string | undefined, dflt: boolean) => (v === undefined ? dflt : v === '1' || v === 'true');
  for (const fc of children(node, 'filterColumn')) {
    const col = rect.c1 + Number(attr(fc, 'colId') ?? 0);
    const [el, ...rest] = fc.children;
    // A column with only its button's settings filters nothing.
    if (!el) continue;
    if (rest.length) return undefined;
    const kind = el.name.replace(/^\w+:/, '');
    if (kind === 'filters') {
      if (children(el, 'dateGroupItem').length) return undefined;
      const values = children(el, 'filter').map((f) => attr(f, 'val') ?? '');
      if (on(attr(el, 'blank'), false)) values.push('');
      columns[col] = values;
    } else if (kind === 'customFilters') {
      const rules = children(el, 'customFilter').map((f) => ({ operator: attr(f, 'operator') ?? 'equal', value: attr(f, 'val') ?? '' }));
      if (!rules.length || rules.length > 2 || rules.some((r) => !FILTER_OPERATORS.has(r.operator))) return undefined;
      conditions[col] = { kind: 'custom', and: on(attr(el, 'and'), false), rules: rules as { operator: FilterOperator; value: string }[] };
    } else if (kind === 'top10') {
      const threshold = Number(attr(el, 'filterVal'));
      conditions[col] = {
        kind: 'top10',
        top: on(attr(el, 'top'), true),
        percent: on(attr(el, 'percent'), false),
        count: Number(attr(el, 'val') ?? 10),
        ...(Number.isFinite(threshold) && attr(el, 'filterVal') !== undefined && { threshold }),
      };
    } else if (kind === 'dynamicFilter' && (attr(el, 'type') === 'aboveAverage' || attr(el, 'type') === 'belowAverage')) {
      const average = Number(attr(el, 'val'));
      conditions[col] = { kind: 'average', above: attr(el, 'type') === 'aboveAverage', ...(Number.isFinite(average) && attr(el, 'val') !== undefined && { average }) };
    } else {
      return undefined;
    }
  }
  return { columns, ...(Object.keys(conditions).length && { conditions }) };
}

/** Whether a read filter picks anything: values in a column, or a condition. */
const picks = (f: Pick<FilterState, 'columns' | 'conditions'> | null | undefined): f is Pick<FilterState, 'columns' | 'conditions'> =>
  !!f && (Object.keys(f.columns).length > 0 || Object.keys(f.conditions ?? {}).length > 0);

function colWidthPx(width: number): number {
  return Math.max(0, Math.round(width * 7));
}

function anchorOf(node: XmlNode | undefined): Anchor | undefined {
  if (!node) return undefined;
  const num = (name: string) => Number(child(node, name)?.text.trim() ?? 0);
  return {
    col: num('col'),
    row: num('row'),
    colOffPx: Math.round(num('colOff') / EMU_PER_PX),
    rowOffPx: Math.round(num('rowOff') / EMU_PER_PX),
  };
}

function readChart(pkg: OoxmlPackage, path: string, ctx: ChartContext): { title?: string; type?: Text; chart?: ChartSpec } {
  const xml = pkg.text(path);
  if (!xml) return {};
  const root = parseXml(xml);
  const chart = child(root, 'chart');
  const titleNode = child(chart, 'title');
  let title = textOf(titleNode).trim();
  if (!title && titleNode) title = descendants(titleNode, 'v').map((v) => v.text).join('').trim();
  if (!title) {
    // Untitled single-series charts are labelled with the series name in Excel.
    const series = descendants(chart, 'ser');
    if (series.length === 1) title = descendants(child(series[0], 'tx'), 'v').map((v) => v.text).join('').trim();
  }
  const plot = child(chart, 'plotArea');
  let type: Text | undefined;
  for (const c of plot?.children ?? []) {
    if (c.name.endsWith('Chart')) {
      // Kinds of chart without a message keep Excel's name for them.
      type = CHART_TYPES[c.name] ? msg(CHART_TYPES[c.name]) : c.name.replace(/Chart$/, '');
      if (c.name === 'barChart' && attr(child(c, 'barDir'), 'val') === 'col') type = msg('core.chartType.column');
      break;
    }
  }
  if (!type) {
    const layout = descendants(root, 'series')[0];
    const id = layout ? attr(layout, 'layoutId') : undefined;
    if (id) type = id.charAt(0).toUpperCase() + id.slice(1);
  }
  return { title: title || undefined, type, chart: plot ? parseChart(root, ctx) : undefined };
}

function readDrawing(pkg: OoxmlPackage, path: string, ctx: ChartContext): SheetObject[] {
  const xml = pkg.text(path);
  if (!xml) return [];
  const rels = new Map(pkg.rels(path).map((r) => [r.id, r]));
  const out: SheetObject[] = [];
  for (const [i, anchor] of parseXml(xml).children.entries()) {
    if (!anchor.name.endsWith('Anchor')) continue;
    const source = { drawing: path, anchor: i };
    const from = anchorOf(child(anchor, 'from'));
    const to = anchorOf(child(anchor, 'to'));
    const ext = child(anchor, 'ext');
    const sizePx = ext
      ? { w: Math.round(Number(attr(ext, 'cx') ?? 0) / EMU_PER_PX), h: Math.round(Number(attr(ext, 'cy') ?? 0) / EMU_PER_PX) }
      : undefined;
    const name = attr(descendants(anchor, 'cNvPr')[0], 'name') ?? '';
    const chartRef = descendants(anchor, 'chart').find((c) => attr(c, 'id'));
    const rel = chartRef ? rels.get(attr(chartRef, 'id') ?? '') : undefined;
    if (rel && (rel.type === 'chart' || rel.type === 'chartEx')) {
      const { title, type, chart } = readChart(pkg, rel.target, ctx);
      // Excel's newer kinds of chart (chartEx: waterfall, treemap…) aren't read.
      out.push({ kind: 'chart', name, title, chartType: type, chart: rel.type === 'chart' ? chart : undefined, from, to, sizePx, source, chartPart: rel.target });
    } else if (descendants(anchor, 'pic').length) {
      out.push({ kind: 'image', name, from, to, sizePx, source });
    } else if (descendants(anchor, 'sp').length || descendants(anchor, 'grpSp').length || descendants(anchor, 'cxnSp').length) {
      out.push({ kind: 'shape', name, from, to, sizePx, source });
    } else if (descendants(anchor, 'graphicFrame').length) {
      out.push({ kind: 'shape', name, from, to, sizePx, source });
    }
  }
  return out;
}

/**
 * The cell metadata indexes (the cm attribute of a cell, counting from 1)
 * that mark a formula as a dynamic array: entries pointing at an XLDAPR
 * block in the metadata part that says the formula is dynamic.
 */
export function dynamicArrayMetadata(xml: string | undefined): Set<number> {
  const out = new Set<number>();
  if (!xml) return out;
  const root = parseXml(xml);
  const type = children(child(root, 'metadataTypes'), 'metadataType').findIndex((t) => attr(t, 'name') === 'XLDAPR') + 1;
  if (!type) return out;
  const future = children(root, 'futureMetadata').find((f) => attr(f, 'name') === 'XLDAPR');
  const dynamic = children(future, 'bk').map((bk) => {
    const props = descendants(bk, 'dynamicArrayProperties')[0];
    return ['1', 'true'].includes(attr(props, 'fDynamic') ?? '');
  });
  children(child(root, 'cellMetadata'), 'bk').forEach((bk, i) => {
    if (children(bk, 'rc').some((rc) => Number(attr(rc, 't')) === type && dynamic[Number(attr(rc, 'v'))])) out.add(i + 1);
  });
  return out;
}

/** Iterative calculation from the workbook's calculation settings, when on: Excel's defaults are 100 passes and a maximum change of 0.001. */
function iterationOf(calcPr: XmlNode | undefined): Iteration | undefined {
  if (!['1', 'true'].includes(attr(calcPr, 'iterate') ?? '')) return undefined;
  const count = Math.trunc(Number(attr(calcPr, 'iterateCount') ?? 100));
  const delta = Number(attr(calcPr, 'iterateDelta') ?? 0.001);
  return { count: count >= 1 && count <= 32767 ? count : 100, delta: Number.isFinite(delta) && delta >= 0 ? delta : 0.001 };
}

/** The error codes a rich value's errorType numbers, from 0. */
const RICH_ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#CALC!'];

/**
 * The errors newer than the file format (#CALC!, #SPILL! and their kind),
 * which Excel saves as #VALUE! for older readers, by value metadata index (a
 * cell's vm attribute, counting from 1): each points through the metadata
 * part at a rich value of the _error structure that names the real one.
 */
export function richErrorMetadata(metadata: string | undefined, values: string | undefined, structures: string | undefined): Map<number, string> {
  const out = new Map<number, string>();
  if (!metadata || !values || !structures) return out;
  const root = parseXml(metadata);
  const type = children(child(root, 'metadataTypes'), 'metadataType').findIndex((t) => attr(t, 'name') === 'XLRICHVALUE') + 1;
  if (!type) return out;
  const future = children(root, 'futureMetadata').find((f) => attr(f, 'name') === 'XLRICHVALUE');
  const rich = children(future, 'bk').map((bk) => Number(attr(descendants(bk, 'rvb')[0], 'i')));
  const kinds = children(parseXml(structures), 's').map((s) => ({ t: attr(s, 't'), keys: children(s, 'k').map((k) => attr(k, 'n')) }));
  const rvs = children(parseXml(values), 'rv').map((rv) => ({ s: Number(attr(rv, 's')), v: children(rv, 'v').map((v) => v.text) }));
  children(child(root, 'valueMetadata'), 'bk').forEach((bk, i) => {
    const rc = children(bk, 'rc').find((x) => Number(attr(x, 't')) === type);
    const rv = rc && rvs[rich[Number(attr(rc, 'v'))]];
    const kind = rv && kinds[rv.s];
    if (!rv || kind?.t !== '_error') return;
    const code = RICH_ERRORS[Number(rv.v[kind.keys.indexOf('errorType')])];
    if (code) out.set(i + 1, code);
  });
  return out;
}

/** Whether a workbook's rich values hold anything but errors: pictures in cells, or linked data types. */
function richValuesBeyondErrors(structures: string | undefined): boolean {
  return !!structures && children(parseXml(structures), 's').some((s) => attr(s, 't') !== '_error');
}

/** Excel 4.0 macro sheets, which hold macros that can run when the file is opened. */
export function isMacroSheet(relType: string | undefined): boolean {
  return relType === 'xlMacrosheet' || relType === 'xlIntlMacrosheet';
}

function countMatches(xml: string | undefined, re: RegExp): number {
  return xml ? (xml.match(re) ?? []).length : 0;
}

/** `cells` counts down the cells read across all sheets; reading stops when it runs out. */
export function readWorkbook(pkg: OoxmlPackage, locale: string, cells?: CellBudget): WorkbookModel {
  const rootRels = pkg.rels('');
  const workbookPath =
    rootRels.find((r) => r.type === 'officeDocument')?.target ?? (pkg.has('xl/workbook.xml') ? 'xl/workbook.xml' : '');
  const wbXml = pkg.text(workbookPath);
  if (!wbXml) throw new TextError(msg('core.open.noWorkbookPart'));
  const wb = parseXml(wbXml);
  const wbRels = pkg.rels(workbookPath);
  const relById = new Map(wbRels.map((r) => [r.id, r]));
  const relOfType = (type: string) => wbRels.filter((r) => r.type === type);
  const people = parsePeople(relOfType('person').map((r) => pkg.text(r.target)).find((x) => x));
  const persons = new Map(people.map((p) => [p.id, p.displayName]));

  const date1904 = ['1', 'true'].includes(attr(child(wb, 'workbookPr'), 'date1904') ?? '');
  const iteration = iterationOf(child(wb, 'calcPr'));
  const metadataXml = relOfType('sheetMetadata').map((r) => pkg.text(r.target)).find((x) => x);
  const dynamic = dynamicArrayMetadata(metadataXml);
  const richPart = (type: string) => relOfType(type).map((r) => pkg.text(r.target)).find((x) => x);
  const richStructures = richPart('rdRichValueStructure');
  const richErrors = richErrorMetadata(metadataXml, richPart('rdRichValue'), richStructures);
  const sst = parseSharedStrings(relOfType('sharedStrings')[0] ? pkg.text(relOfType('sharedStrings')[0].target) : undefined);
  const styles = new StyleBook(
    relOfType('styles')[0] ? pkg.text(relOfType('styles')[0].target) : undefined,
    relOfType('theme')[0] ? pkg.text(relOfType('theme')[0].target) : undefined,
    locale,
    relOfType('FeaturePropertyBag').map((r) => pkg.text(r.target)).find((x) => x),
  );

  const findings: Finding[] = [];
  const sheetNodes = children(child(wb, 'sheets'), 'sheet');
  const sheetNames = sheetNodes.map((s) => attr(s, 'name') ?? '');
  const sheetIndexByName = new Map(sheetNames.map((n, i) => [n.toLowerCase(), i]));

  const names: DefinedName[] = children(child(wb, 'definedNames'), 'definedName').map((d, fileIndex) => {
    const local = attr(d, 'localSheetId');
    const comment = attr(d, 'comment');
    return {
      name: attr(d, 'name') ?? '',
      formula: d.text,
      localSheet: local !== undefined ? Number(local) : undefined,
      ...(['1', 'true'].includes(attr(d, 'hidden') ?? '') && { hidden: true }),
      ...(comment && { comment }),
      fileIndex,
    };
  });

  const chartContext: ChartContext = { theme: styles.theme, sheetIndex: sheetIndexByName, names };

  // Pivot caches, keyed by their definition part path, to describe each pivot's source.
  const pivotSources = new Map<string, Text>();
  for (const r of relOfType('pivotCacheDefinition')) {
    const def = pkg.text(r.target);
    if (!def) continue;
    const ws = descendants(parseXml(def), 'worksheetSource')[0];
    const src = ws ? [attr(ws, 'sheet'), attr(ws, 'ref') ?? attr(ws, 'name')].filter(Boolean).join('!') : msg('core.pivot.externalSource');
    pivotSources.set(r.target, src);
  }

  const sheets: SheetModel[] = sheetNodes.map((node, index): SheetModel => {
    const name = sheetNames[index];
    const rel = relById.get(attr(node, 'id') ?? '');
    const state = (attr(node, 'state') as SheetInfo['state']) ?? 'visible';
    const path = rel?.target ?? '';
    const info: SheetInfo = {
      index,
      name,
      kind: rel?.type === 'worksheet' ? 'worksheet' : rel?.type === 'chartsheet' ? 'chartsheet' : 'other',
      state,
      maxRow: -1,
      maxCol: -1,
      defaultColWidthPx: 64,
      defaultRowHeightPx: 20,
      colWidths: {},
      rowHeights: {},
      hiddenCols: [],
      hiddenRows: [],
      frozenCols: 0,
      frozenRows: 0,
      rowStyles: {},
      colStyles: [],
      merges: [],
      listValidations: [],
      ruleValidations: [],
      blocked: [],
      objects: [],
      notes: [],
      tables: [],
      conditionalFormats: [],
      links: [],
      protected: false,
      showGridLines: true,
    };
    if (state !== 'visible') {
      findings.push({
        kind: 'hiddenSheet',
        severity: 'info',
        title: msg('core.finding.hiddenSheet.title', { sheet: name }),
        detail: msg('core.finding.hiddenSheet.detail'),
        sheet: name,
      });
    }
    const sheetRels = path ? pkg.rels(path) : [];

    if (info.kind === 'chartsheet') {
      const drawing = sheetRels.find((r) => r.type === 'drawing');
      const objects = drawing ? readDrawing(pkg, drawing.target, chartContext) : [];
      info.objects = objects;
      const chart = objects.find((o) => o.kind === 'chart');
      findings.push({
        kind: 'chartsheet',
        severity: 'warn',
        title: msg('core.finding.chartsheet.title', { sheet: name }),
        detail: chart?.chartType
          ? chart.title
            ? msg('core.finding.chartsheet.detailTypedTitled', { type: chart.chartType, title: chart.title })
            : msg('core.finding.chartsheet.detailTyped', { type: chart.chartType })
          : chart?.title
            ? msg('core.finding.chartsheet.detailTitled', { title: chart.title })
            : msg('core.finding.chartsheet.detail'),
        sheet: name,
      });
      return { info, path, colStyles: [] };
    }
    if (info.kind !== 'worksheet') {
      const macroSheet = isMacroSheet(rel?.type);
      findings.push({
        kind: 'macro',
        severity: 'warn',
        title: msg(macroSheet ? 'core.finding.macroSheet.title' : 'core.finding.unsupportedSheet.title', { sheet: name }),
        detail: msg(macroSheet ? 'core.finding.macroSheet.detail' : 'core.finding.unsupportedSheet.detail'),
        sheet: name,
      });
      return { info, path, colStyles: [] };
    }

    const xml = pkg.text(path) ?? '';
    const scan = scanSheet(xml, sst.strings, sst.rich, cells, dynamic, richErrors);
    info.maxRow = scan.maxRow;
    info.maxCol = scan.maxCol;

    // Everything outside <sheetData> is small; parse it as one document.
    const meta =
      scan.sheetDataStart >= 0 ? parseXml(xml.slice(0, scan.sheetDataStart) + xml.slice(scan.sheetDataEnd)) : parseXml(xml);
    const fmt = child(meta, 'sheetFormatPr');
    if (attr(fmt, 'defaultColWidth')) info.defaultColWidthPx = colWidthPx(Number(attr(fmt, 'defaultColWidth')));
    if (attr(fmt, 'defaultRowHeight')) info.defaultRowHeightPx = Math.round((Number(attr(fmt, 'defaultRowHeight')) * 4) / 3);
    for (const [row, pt] of scan.rowHeightsPt) info.rowHeights[row] = Math.round((pt * 4) / 3);
    if (scan.fittedRows.length) info.fittedRows = Object.fromEntries(scan.fittedRows.map((r) => [r, true as const]));
    info.hiddenRows = scan.hiddenRows;
    for (const [row, style] of scan.rowStyles) info.rowStyles[row] = style;

    const colStyles = info.colStyles;
    const colLimit = Math.max(info.maxCol, 25) + COLUMN_PADDING;
    const colLevels: Record<number, number> = {};
    const collapsedCols: number[] = [];
    const yes = (v: string | undefined) => v === '1' || v === 'true';
    for (const c of children(child(meta, 'cols'), 'col')) {
      const min = Number(attr(c, 'min')) - 1;
      const max = Number(attr(c, 'max')) - 1;
      const width = attr(c, 'width');
      const hidden = yes(attr(c, 'hidden'));
      const style = Number(attr(c, 'style') ?? 0);
      const level = Math.min(7, Number(attr(c, 'outlineLevel') ?? 0));
      const collapsed = yes(attr(c, 'collapsed'));
      if (style) colStyles.push({ min, max, style });
      for (let col = min; col <= Math.min(max, colLimit); col++) {
        if (width !== undefined) info.colWidths[col] = colWidthPx(Number(width));
        if (hidden) info.hiddenCols.push(col);
        if (level > 0) colLevels[col] = level;
        if (collapsed) collapsedCols.push(col);
      }
    }
    // Groups' summaries are below and right of them unless <outlinePr> says otherwise.
    const outlinePr = child(child(meta, 'sheetPr'), 'outlinePr');
    const after = (name: string) => !['0', 'false'].includes(attr(outlinePr, name) ?? '');
    if (outlinePr || scan.rowLevels.size || scan.collapsedRows.length || collapsedCols.length || Object.keys(colLevels).length) {
      info.outline = {
        rows: { levels: Object.fromEntries(scan.rowLevels), collapsed: scan.collapsedRows, summaryAfter: after('summaryBelow') },
        cols: { levels: colLevels, collapsed: collapsedCols, summaryAfter: after('summaryRight') },
      };
    }

    const view = child(child(meta, 'sheetViews'), 'sheetView');
    if (attr(view, 'showGridLines') === '0' || attr(view, 'showGridLines') === 'false') info.showGridLines = false;
    const pane = child(view, 'pane');
    if (pane && /frozen/i.test(attr(pane, 'state') ?? '')) {
      info.frozenCols = Number(attr(pane, 'xSplit') ?? 0);
      info.frozenRows = Number(attr(pane, 'ySplit') ?? 0);
    }
    const tab = child(child(meta, 'sheetPr'), 'tabColor');
    if (tab && attr(tab, 'rgb')) info.tabColor = '#' + (attr(tab, 'rgb') ?? '').slice(-6);

    info.merges = children(child(meta, 'mergeCells'), 'mergeCell')
      .map((m) => parseRange(attr(m, 'ref') ?? ''))
      .filter((r): r is Rect => !!r);
    for (const m of info.merges) {
      info.maxRow = Math.max(info.maxRow, m.r2);
      info.maxCol = Math.max(info.maxCol, m.c2);
    }
    const tallMerges = info.merges.filter((m) => m.r2 > m.r1).length;
    if (tallMerges) {
      findings.push({
        kind: 'mergedCells',
        severity: 'info',
        title: msg('core.finding.mergedCells.title'),
        detail: msg('core.finding.mergedCells.detail', { count: tallMerges }),
        sheet: name,
      });
    }

    const protection = child(meta, 'sheetProtection');
    if (protection && ['1', 'true'].includes(attr(protection, 'sheet') ?? '')) {
      info.protected = true;
      info.objectsLocked = ['1', 'true'].includes(attr(protection, 'objects') ?? '');
      // These say what the protection locks, so "0" is what allows it.
      const allows = (name: string) => ['0', 'false'].includes(attr(protection, name) ?? '');
      info.formatAllowed = { rows: allows('formatRows'), cols: allows('formatColumns') };
      findings.push({
        kind: 'protection',
        severity: 'info',
        title: msg('core.finding.protectedSheet.title'),
        detail: msg('core.finding.protectedSheet.detail'),
        sheet: name,
      });
    }

    // Dynamic arrays are locked too until the engine shows it calculates them, when they are let go.
    const blocked: BlockedRegion[] = scan.arrayRanges.map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2, kind: 'array', reason: msg('core.reason.arrayFormula') }));
    const legacy = scan.arrayRanges.filter((a) => !a.dynamic);
    if (legacy.length) {
      findings.push({
        kind: 'arrayFormula',
        severity: 'info',
        title: msg('core.finding.arrayFormulas.title'),
        detail: msg('core.finding.arrayFormulas.detail', { count: legacy.length }),
        sheet: name,
        where: legacy.slice(0, 5).map(rangeToA1).join(', '),
      });
    }

    info.conditionalFormats = parseConditionalFormats(meta, styles);
    if (info.conditionalFormats.length) {
      findings.push({
        kind: 'conditionalFormat',
        severity: 'info',
        title: msg('core.finding.conditionalFormat.title'),
        detail: msg('core.finding.conditionalFormat.detail', { count: info.conditionalFormats.length }),
        sheet: name,
      });
    }
    const validations = parseValidations(meta, { sheetIndex: sheetIndexByName, ownSheet: index, names });
    info.listValidations = validations.lists;
    info.ruleValidations = validations.rules;
    if (validations.other) info.validationsKept = true;
    // "Any value" rules only show a message; like rules the editor can't read, nothing is checked.
    const checked = validations.rules.filter((r) => r.kind !== 'any').length;
    const unchecked = validations.rules.length - checked + validations.other;
    const dvCount = validations.lists.length + checked + unchecked;
    if (dvCount) {
      const parts: Message[] = [];
      if (validations.lists.length) parts.push(msg('core.finding.dataValidation.lists', { count: validations.lists.length }));
      if (checked) parts.push(msg('core.finding.dataValidation.rules', { count: checked }));
      if (unchecked) parts.push(msg('core.finding.dataValidation.other', { count: unchecked }));
      findings.push({
        kind: 'dataValidation',
        severity: 'info',
        title: msg('core.finding.dataValidation.title'),
        detail: msg('core.finding.dataValidation.detail', { rules: joinTexts(parts, 'core.list.clauses') }),
        sheet: name,
      });
    }
    info.links = readHyperlinks(meta, sheetRels);
    if (info.links.length) {
      findings.push({
        kind: 'hyperlink',
        severity: 'info',
        title: msg('links.finding.title'),
        detail: msg('links.finding.detail', { count: info.links.length }),
        sheet: name,
      });
    }
    const sparkCount = countMatches(xml.slice(scan.sheetDataEnd), /<x14:sparklineGroup\b/g);
    if (sparkCount) {
      findings.push({
        kind: 'sparkline',
        severity: 'info',
        title: msg('core.finding.sparklines.title'),
        detail: msg('core.finding.sparklines.detail', { count: sparkCount }),
        sheet: name,
      });
    }
    if (child(meta, 'autoFilter')) {
      info.autoFilter = parseRange(attr(child(meta, 'autoFilter'), 'ref') ?? '') ?? undefined;
      const read = info.autoFilter && filterColumns(child(meta, 'autoFilter'), info.autoFilter);
      if (info.autoFilter && picks(read)) info.filter = { rect: info.autoFilter, ...read };
      if (info.autoFilter && !read) info.filterKept = true;
      findings.push({
        kind: 'autoFilter',
        severity: 'info',
        title: msg('core.finding.filter.title'),
        detail: msg('core.finding.filter.detail'),
        sheet: name,
      });
    }
    if (scan.richTextCells) {
      findings.push({
        kind: 'richText',
        severity: 'info',
        title: msg('core.finding.richText.title'),
        detail: msg('core.finding.richText.detail', { count: scan.richTextCells }),
        sheet: name,
      });
    }

    const notes: CellNote[] = [];
    const threads: CellNote[] = [];
    for (const r of sheetRels) {
      if (r.type === 'drawing') {
        info.objects.push(...readDrawing(pkg, r.target, chartContext));
      } else if (r.type === 'pivotTable') {
        const pxml = pkg.text(r.target);
        if (!pxml) continue;
        const pt = parseXml(pxml);
        const pname = attr(pt, 'name');
        const shownName: Text = pname ?? msg('core.pivot.defaultName');
        const ref = attr(child(pt, 'location'), 'ref') ?? '';
        const range = parseRange(ref);
        const cacheRel = pkg.rels(r.target).find((x) => x.type === 'pivotCacheDefinition');
        const source = cacheRel ? pivotSources.get(cacheRel.target) : undefined;
        if (range) {
          blocked.push({ ...range, kind: 'pivot', reason: msg('core.reason.pivot', { name: shownName }) });
          info.objects.push({ kind: 'pivot', name: pname ?? '', from: { row: range.r1, col: range.c1, rowOffPx: 0, colOffPx: 0 } });
          info.maxRow = Math.max(info.maxRow, range.r2);
          info.maxCol = Math.max(info.maxCol, range.c2);
        }
        findings.push({
          kind: 'pivot',
          severity: 'warn',
          title: msg('core.finding.pivot.title', { name: shownName }),
          detail: msg('core.finding.pivot.detail', { source: source ?? msg('core.pivot.otherRange') }),
          sheet: name,
          where: ref,
        });
      } else if (r.type === 'table') {
        const txml = pkg.text(r.target);
        if (!txml) continue;
        const t = parseXml(txml);
        const tname = attr(t, 'displayName') ?? attr(t, 'name') ?? 'Table';
        const ref = attr(t, 'ref') ?? '';
        const range = parseRange(ref);
        const headerRows = Number(attr(t, 'headerRowCount') ?? 1);
        if (range) {
          const si = child(t, 'tableStyleInfo');
          const style = attr(si, 'name');
          const on = (name: string) => attr(si, name) === '1' || attr(si, name) === 'true';
          info.tables.push({
            name: tname,
            rect: range,
            headerRows,
            totalsRows: Number(attr(t, 'totalsRowCount') ?? 0),
            filterButtons: headerRows > 0 && !!child(t, 'autoFilter'),
            columns: children(child(t, 'tableColumns'), 'tableColumn').map((c) => attr(c, 'name') ?? ''),
            calculated: children(child(t, 'tableColumns'), 'tableColumn').map((c) => child(c, 'calculatedColumnFormula')?.text || undefined),
            path: r.target,
            style,
            showFirstColumn: on('showFirstColumn'),
            showLastColumn: on('showLastColumn'),
            showRowStripes: on('showRowStripes'),
            showColumnStripes: on('showColumnStripes'),
            look: style ? (styles.tableStyles.get(style) ?? builtinTableLook(style, styles.theme)) : undefined,
          });
          // A table's own filter, if the sheet has no other with values picked.
          const af = child(t, 'autoFilter');
          const afRect = af && parseRange(attr(af, 'ref') ?? '');
          const read = afRect && filterColumns(af, afRect);
          if (afRect && picks(read) && !info.filter) info.filter = { rect: afRect, ...read, table: tname };
          if (afRect && !read) info.tables[info.tables.length - 1].filterKept = true;
        }
        if (range && headerRows > 0) {
          blocked.push({
            r1: range.r1,
            c1: range.c1,
            r2: range.r1,
            c2: range.c2,
            kind: 'tableHeader',
            reason: msg('core.reason.tableHeader', { name: tname }),
          });
        }
        findings.push({
          kind: 'table',
          severity: 'info',
          title: msg('core.finding.table.title', { name: tname }),
          detail: msg('core.finding.table.detail'),
          sheet: name,
          where: ref,
        });
      } else if (r.type === 'comments') {
        const xml = pkg.text(r.target);
        if (xml) notes.push(...parseNotes(xml));
      } else if (r.type === 'threadedComment') {
        const xml = pkg.text(r.target);
        if (xml) threads.push(...parseThreads(xml, persons));
      } else if (r.type === 'ctrlProp') {
        info.objects.push({ kind: 'control', name: 'Form control' });
      } else if (r.type === 'slicer' || r.type === 'timeline') {
        findings.push({
          kind: 'slicer',
          severity: 'info',
          title: msg(r.type === 'slicer' ? 'core.finding.slicers.title' : 'core.finding.timelines.title'),
          detail: msg('core.finding.slicers.detail'),
          sheet: name,
        });
      }
    }
    info.notes = mergeNotes(notes, threads);
    if (info.notes.length) {
      const threadCount = threads.length;
      const noteCount = info.notes.length - threadCount;
      const notesText = msg('core.finding.comments.notes', { count: noteCount });
      const threadsText = msg('core.finding.comments.threads', { count: threadCount });
      const what = noteCount && threadCount ? msg('core.finding.comments.notesAndThreads', { notes: notesText, threads: threadsText }) : noteCount ? notesText : threadsText;
      findings.push({
        kind: 'comment',
        severity: 'info',
        title: msg('core.finding.comments.title'),
        detail: msg('core.finding.comments.detail', { what }),
        sheet: name,
      });
    }

    const charts = info.objects.filter((o) => o.kind === 'chart');
    for (const c of charts) {
      const at = c.from ? toA1(c.from.row, c.from.col) : undefined;
      findings.push({
        kind: 'chart',
        severity: 'warn',
        title: c.title ? msg('core.finding.chart.titleNamed', { title: c.title }) : msg('core.finding.chart.title'),
        detail: c.chartType ? msg('core.finding.chart.detailTyped', { type: c.chartType }) : msg('core.finding.chart.detail'),
        sheet: name,
        where: at,
      });
    }
    const images = info.objects.filter((o) => o.kind === 'image').length;
    if (images) {
      findings.push({
        kind: 'image',
        severity: 'info',
        title: msg('core.finding.pictures.title'),
        detail: msg('core.finding.pictures.detail', { count: images }),
        sheet: name,
      });
    }
    const shapes = info.objects.filter((o) => o.kind === 'shape').length;
    if (shapes) {
      findings.push({
        kind: 'shape',
        severity: 'info',
        title: msg('core.finding.shapes.title'),
        detail: msg('core.finding.shapes.detail', { count: shapes }),
        sheet: name,
      });
    }
    const controls = info.objects.filter((o) => o.kind === 'control').length;
    if (controls) {
      findings.push({
        kind: 'control',
        severity: 'info',
        title: msg('core.finding.controls.title'),
        detail: msg('core.finding.controls.detail', { count: controls }),
        sheet: name,
      });
    }

    info.blocked = blocked;
    // Big sheets' text is several times the size of their parsed cells, so it
    // isn't kept: saving and restructuring decode it again from the package.
    pkg.forgetText(path);
    const sheetModel: SheetModel = { info, path, scan, colStyles };
    Object.defineProperty(sheetModel, 'xml', { get: () => pkg.textOnce(path), enumerable: true });
    return sheetModel;
  });

  // Workbook-level features.
  const vba = relOfType('vbaProject').length > 0 || pkg.has('xl/vbaProject.bin');
  const macroSheets = wbRels.some((r) => isMacroSheet(r.type));
  // Controls and embedded objects can run code in Excel, macros or not.
  const controls = pkg.list().filter((p) => /^xl\/activeX\/activeX\d*\.xml$/.test(p)).length;
  const objects = pkg.list().filter((p) => p.startsWith('xl/embeddings/')).length;
  if (controls || objects) {
    findings.push({
      kind: 'activeContent',
      severity: 'warn',
      title: msg('core.finding.activeContent.title'),
      detail: msg('core.finding.activeContent.detail', { controls, objects }),
    });
  }
  if (vba) {
    findings.push({
      kind: 'macro',
      severity: 'warn',
      title: msg('core.finding.macros.title'),
      detail: msg('core.finding.macros.detail'),
    });
  }
  const ext: Relationship[] = relOfType('externalLink');
  if (ext.length) {
    findings.push({
      kind: 'externalLink',
      severity: 'warn',
      title: msg('core.finding.externalLinks.title'),
      detail: msg('core.finding.externalLinks.detail', { count: ext.length }),
    });
  }
  if (relOfType('connections').length || pkg.list().some((p) => p.startsWith('xl/queryTables/'))) {
    findings.push({
      kind: 'connection',
      severity: 'info',
      title: msg('core.finding.connections.title'),
      detail: msg('core.finding.connections.detail'),
    });
  }
  if (relOfType('slicerCache').length || relOfType('timelineCache').length) {
    if (!findings.some((f) => f.kind === 'slicer')) {
      findings.push({ kind: 'slicer', severity: 'info', title: msg('core.finding.slicers.title'), detail: msg('core.finding.slicers.detail') });
    }
  }
  // Rich values that only name newer errors aren't pictures.
  if (pkg.list().some((p) => p.startsWith('xl/richData/')) && (!richStructures || richValuesBeyondErrors(richStructures))) {
    findings.push({
      kind: 'cellImage',
      severity: 'info',
      title: msg('core.finding.cellImages.title'),
      detail: msg('core.finding.cellImages.detail'),
    });
  }
  if (child(wb, 'workbookProtection')) {
    findings.push({
      kind: 'protection',
      severity: 'info',
      title: msg('core.finding.workbookProtection.title'),
      detail: msg('core.finding.workbookProtection.detail'),
    });
  }

  return {
    workbookPath,
    sheets,
    styles,
    names,
    date1904,
    ...(iteration && { iteration }),
    findings,
    hasMacros: vba || macroSheets,
    hasActiveContent: controls + objects > 0,
    properties: readCustomProperties(pkg),
    people,
    structureLocked: ['1', 'true'].includes(attr(child(wb, 'workbookProtection'), 'lockStructure') ?? ''),
    chartContext,
  };
}

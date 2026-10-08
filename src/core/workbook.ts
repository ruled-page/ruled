import { msg, TextError, type Message, type MessageKey, type Text } from '../i18n/message';
import { parseChart, type ChartContext, type ChartSpec } from './charts';
import { parseRange, rangeToA1, toA1 } from './refs';
import { parseSharedStrings, scanSheet, type CellBudget, type ScanResult } from './sheetScanner';
import { parseConditionalFormats } from './conditionalFormat';
import { readCustomProperties } from './properties';
import { StyleBook } from './styles';
import { builtinTableLook } from './tableStyles';
import { mergeNotes, parseNotes, parsePersons, parseThreads } from './comments';
import type { Anchor, BlockedRegion, CellNote, Finding, Rect, SheetInfo, SheetObject } from './types';
import { parseValidations } from './validations';
import { attr, child, children, descendants, parseXml, textOf, type XmlNode } from './xml';
import type { OoxmlPackage, Relationship } from './zip';

export interface DefinedName {
  name: string;
  formula: string;
  /** Index of the sheet the name is local to, if any. */
  localSheet?: number;
}

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
  findings: Finding[];
  /** The workbook carries a VBA project. */
  /** A VBA project or Excel 4.0 macro sheets. */
  hasMacros: boolean;
  /** ActiveX controls or embedded (OLE) objects. */
  hasActiveContent: boolean;
  /** Custom document properties. */
  properties: Record<string, string>;
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
  for (const anchor of parseXml(xml).children) {
    if (!anchor.name.endsWith('Anchor')) continue;
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
      out.push({ kind: 'chart', name, title, chartType: type, chart: rel.type === 'chart' ? chart : undefined, from, to, sizePx });
    } else if (descendants(anchor, 'pic').length) {
      out.push({ kind: 'image', name, from, to, sizePx });
    } else if (descendants(anchor, 'sp').length || descendants(anchor, 'grpSp').length || descendants(anchor, 'cxnSp').length) {
      out.push({ kind: 'shape', name, from, to, sizePx });
    } else if (descendants(anchor, 'graphicFrame').length) {
      out.push({ kind: 'shape', name, from, to, sizePx });
    }
  }
  return out;
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
  const persons = parsePersons(relOfType('person').map((r) => pkg.text(r.target)).find((x) => x));

  const date1904 = ['1', 'true'].includes(attr(child(wb, 'workbookPr'), 'date1904') ?? '');
  const sst = parseSharedStrings(relOfType('sharedStrings')[0] ? pkg.text(relOfType('sharedStrings')[0].target) : undefined);
  const styles = new StyleBook(
    relOfType('styles')[0] ? pkg.text(relOfType('styles')[0].target) : undefined,
    relOfType('theme')[0] ? pkg.text(relOfType('theme')[0].target) : undefined,
    locale,
  );

  const findings: Finding[] = [];
  const sheetNodes = children(child(wb, 'sheets'), 'sheet');
  const sheetNames = sheetNodes.map((s) => attr(s, 'name') ?? '');
  const sheetIndexByName = new Map(sheetNames.map((n, i) => [n.toLowerCase(), i]));

  const names: DefinedName[] = children(child(wb, 'definedNames'), 'definedName').map((d) => {
    const local = attr(d, 'localSheetId');
    return { name: attr(d, 'name') ?? '', formula: d.text, localSheet: local !== undefined ? Number(local) : undefined };
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
    const scan = scanSheet(xml, sst.strings, sst.rich, cells);
    info.maxRow = scan.maxRow;
    info.maxCol = scan.maxCol;

    // Everything outside <sheetData> is small; parse it as one document.
    const meta =
      scan.sheetDataStart >= 0 ? parseXml(xml.slice(0, scan.sheetDataStart) + xml.slice(scan.sheetDataEnd)) : parseXml(xml);
    const fmt = child(meta, 'sheetFormatPr');
    if (attr(fmt, 'defaultColWidth')) info.defaultColWidthPx = colWidthPx(Number(attr(fmt, 'defaultColWidth')));
    if (attr(fmt, 'defaultRowHeight')) info.defaultRowHeightPx = Math.round((Number(attr(fmt, 'defaultRowHeight')) * 4) / 3);
    for (const [row, pt] of scan.rowHeightsPt) info.rowHeights[row] = Math.round((pt * 4) / 3);
    info.hiddenRows = scan.hiddenRows;
    for (const [row, style] of scan.rowStyles) info.rowStyles[row] = style;

    const colStyles = info.colStyles;
    const colLimit = Math.max(info.maxCol, 25) + COLUMN_PADDING;
    for (const c of children(child(meta, 'cols'), 'col')) {
      const min = Number(attr(c, 'min')) - 1;
      const max = Number(attr(c, 'max')) - 1;
      const width = attr(c, 'width');
      const hidden = attr(c, 'hidden') === '1' || attr(c, 'hidden') === 'true';
      const style = Number(attr(c, 'style') ?? 0);
      if (style) colStyles.push({ min, max, style });
      for (let col = min; col <= Math.min(max, colLimit); col++) {
        if (width !== undefined) info.colWidths[col] = colWidthPx(Number(width));
        if (hidden) info.hiddenCols.push(col);
      }
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
      findings.push({
        kind: 'protection',
        severity: 'info',
        title: msg('core.finding.protectedSheet.title'),
        detail: msg('core.finding.protectedSheet.detail'),
        sheet: name,
      });
    }

    const blocked: BlockedRegion[] = [];
    for (const a of scan.arrayRanges) {
      blocked.push({ ...a, kind: 'array', reason: msg('core.reason.arrayFormula') });
    }
    if (scan.arrayRanges.length) {
      findings.push({
        kind: 'arrayFormula',
        severity: 'info',
        title: msg('core.finding.arrayFormulas.title'),
        detail: msg('core.finding.arrayFormulas.detail', { count: scan.arrayRanges.length }),
        sheet: name,
        where: scan.arrayRanges.slice(0, 5).map(rangeToA1).join(', '),
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
    const dvCount = validations.lists.length + validations.rules.length + validations.other;
    if (dvCount) {
      const parts: Message[] = [];
      if (validations.lists.length) parts.push(msg('core.finding.dataValidation.lists', { count: validations.lists.length }));
      if (validations.rules.length) parts.push(msg('core.finding.dataValidation.rules', { count: validations.rules.length }));
      if (validations.other) parts.push(msg('core.finding.dataValidation.other', { count: validations.other }));
      findings.push({
        kind: 'dataValidation',
        severity: 'info',
        title: msg('core.finding.dataValidation.title'),
        detail: msg('core.finding.dataValidation.detail', { rules: joinTexts(parts, 'core.list.clauses') }),
        sheet: name,
      });
    }
    const linkCount = children(child(meta, 'hyperlinks'), 'hyperlink').length;
    if (linkCount) {
      findings.push({
        kind: 'hyperlink',
        severity: 'info',
        title: msg('core.finding.hyperlinks.title'),
        detail: msg('core.finding.hyperlinks.detail', { count: linkCount }),
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
  if (pkg.list().some((p) => p.startsWith('xl/richData/'))) {
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
    findings,
    hasMacros: vba || macroSheets,
    hasActiveContent: controls + objects > 0,
    properties: readCustomProperties(pkg),
  };
}

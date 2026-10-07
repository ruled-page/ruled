import { parseRange, rangeToA1, toA1 } from './refs';
import { parseSharedStrings, scanSheet, type ScanResult } from './sheetScanner';
import { StyleBook } from './styles';
import type { Anchor, BlockedRegion, Finding, Rect, SheetInfo, SheetObject } from './types';
import { parseListValidations } from './validations';
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
  xml?: string;
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
}

const EMU_PER_PX = 9525;
/** Hidden-column and width ranges are only expanded this far past the used range. */
const COLUMN_PADDING = 64;

const CHART_TYPE_LABELS: Record<string, string> = {
  barChart: 'Bar',
  bar3DChart: '3-D bar',
  lineChart: 'Line',
  line3DChart: '3-D line',
  pieChart: 'Pie',
  pie3DChart: '3-D pie',
  doughnutChart: 'Doughnut',
  areaChart: 'Area',
  area3DChart: '3-D area',
  scatterChart: 'Scatter',
  bubbleChart: 'Bubble',
  radarChart: 'Radar',
  stockChart: 'Stock',
  surfaceChart: 'Surface',
  surface3DChart: '3-D surface',
  ofPieChart: 'Pie of pie',
};

function plural(n: number, one: string, many = one + 's'): string {
  return `${n.toLocaleString()} ${n === 1 ? one : many}`;
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

function readChart(pkg: OoxmlPackage, path: string): { title?: string; type?: string } {
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
  let type: string | undefined;
  for (const c of plot?.children ?? []) {
    if (c.name.endsWith('Chart')) {
      type = CHART_TYPE_LABELS[c.name] ?? c.name.replace(/Chart$/, '');
      if (c.name === 'barChart' && attr(child(c, 'barDir'), 'val') === 'col') type = 'Column';
      break;
    }
  }
  if (!type) {
    const layout = descendants(root, 'series')[0];
    const id = layout ? attr(layout, 'layoutId') : undefined;
    if (id) type = id.charAt(0).toUpperCase() + id.slice(1);
  }
  return { title: title || undefined, type };
}

function readDrawing(pkg: OoxmlPackage, path: string): SheetObject[] {
  const xml = pkg.text(path);
  if (!xml) return [];
  const rels = new Map(pkg.rels(path).map((r) => [r.id, r]));
  const out: SheetObject[] = [];
  for (const anchor of parseXml(xml).children) {
    if (!/Anchor$/.test(anchor.name)) continue;
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
      const { title, type } = readChart(pkg, rel.target);
      out.push({ kind: 'chart', name, title, chartType: type, from, to, sizePx });
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

function countMatches(xml: string | undefined, re: RegExp): number {
  return xml ? (xml.match(re) ?? []).length : 0;
}

export function readWorkbook(pkg: OoxmlPackage, locale: string): WorkbookModel {
  const rootRels = pkg.rels('');
  const workbookPath =
    rootRels.find((r) => r.type === 'officeDocument')?.target ?? (pkg.has('xl/workbook.xml') ? 'xl/workbook.xml' : '');
  const wbXml = pkg.text(workbookPath);
  if (!wbXml) throw new Error('The workbook part is missing from this file.');
  const wb = parseXml(wbXml);
  const wbRels = pkg.rels(workbookPath);
  const relById = new Map(wbRels.map((r) => [r.id, r]));
  const relOfType = (type: string) => wbRels.filter((r) => r.type === type);

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

  // Pivot caches, keyed by their definition part path, to describe each pivot's source.
  const pivotSources = new Map<string, string>();
  for (const r of relOfType('pivotCacheDefinition')) {
    const def = pkg.text(r.target);
    if (!def) continue;
    const ws = descendants(parseXml(def), 'worksheetSource')[0];
    const src = ws ? [attr(ws, 'sheet'), attr(ws, 'ref') ?? attr(ws, 'name')].filter(Boolean).join('!') : 'an external source';
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
      merges: [],
      listValidations: [],
      blocked: [],
      objects: [],
      protected: false,
      showGridLines: true,
    };
    if (state !== 'visible') {
      findings.push({
        kind: 'hiddenSheet',
        severity: 'info',
        title: `Hidden sheet "${name}"`,
        detail: 'Kept in the file but not shown, as in Excel.',
        sheet: name,
      });
    }
    const sheetRels = path ? pkg.rels(path) : [];

    if (info.kind === 'chartsheet') {
      const drawing = sheetRels.find((r) => r.type === 'drawing');
      const objects = drawing ? readDrawing(pkg, drawing.target) : [];
      info.objects = objects;
      const chart = objects.find((o) => o.kind === 'chart');
      findings.push({
        kind: 'chartsheet',
        severity: 'warn',
        title: `Chart sheet "${name}"`,
        detail: `${chart?.chartType ? chart.chartType + ' chart' : 'A chart'}${chart?.title ? ` titled "${chart.title}"` : ''} that fills a whole sheet. Kept in the file and redrawn by Excel from the edited data; not shown in this editor.`,
        sheet: name,
      });
      return { info, path, colStyles: [] };
    }
    if (info.kind !== 'worksheet') {
      findings.push({
        kind: 'macro',
        severity: 'warn',
        title: `Unsupported sheet "${name}"`,
        detail: 'A dialog or macro sheet from older versions of Excel. Kept in the file; not shown in this editor.',
        sheet: name,
      });
      return { info, path, colStyles: [] };
    }

    const xml = pkg.text(path) ?? '';
    const scan = scanSheet(xml, sst.strings, sst.rich);
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

    const colStyles: SheetModel['colStyles'] = [];
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
        title: 'Merged cells spanning rows',
        detail: `${plural(tallMerges, 'merged range')} span more than one row. Their value is shown in the top-left cell only; the merge is kept in the file.`,
        sheet: name,
      });
    }
    if (info.frozenRows > 0) {
      findings.push({
        kind: 'frozenRows',
        severity: 'info',
        title: 'Frozen rows',
        detail: `Excel freezes the top ${plural(info.frozenRows, 'row')} on this sheet. This editor can only freeze columns; the setting is kept in the file.`,
        sheet: name,
      });
    }

    const protection = child(meta, 'sheetProtection');
    if (protection && ['1', 'true'].includes(attr(protection, 'sheet') ?? '')) {
      info.protected = true;
      findings.push({
        kind: 'protection',
        severity: 'info',
        title: 'Protected sheet',
        detail: 'This sheet is protected in Excel, so it is read-only here.',
        sheet: name,
      });
    }

    const blocked: BlockedRegion[] = [];
    for (const a of scan.arrayRanges) {
      blocked.push({ ...a, kind: 'array', reason: 'This cell is part of an array formula. Edit it in Excel.' });
    }
    if (scan.arrayRanges.length) {
      findings.push({
        kind: 'arrayFormula',
        severity: 'info',
        title: 'Array formulas',
        detail: `${plural(scan.arrayRanges.length, 'array or spill formula')} cover more than one cell. They keep Excel's last values, Excel recalculates them on open, and their cells can't be edited here.`,
        sheet: name,
        where: scan.arrayRanges.slice(0, 5).map(rangeToA1).join(', '),
      });
    }

    const cfCount =
      children(meta, 'conditionalFormatting').length + countMatches(xml.slice(scan.sheetDataEnd), /<x14:conditionalFormatting\b/g);
    if (cfCount) {
      findings.push({
        kind: 'conditionalFormat',
        severity: 'info',
        title: 'Conditional formatting',
        detail: `${plural(cfCount, 'rule set')}. Kept in the file and applied by Excel; not shown in this editor.`,
        sheet: name,
      });
    }
    const validations = parseListValidations(meta, { sheetIndex: sheetIndexByName, ownSheet: index, names });
    info.listValidations = validations.lists;
    const dvCount = validations.lists.length + validations.other;
    if (dvCount) {
      const unresolved = validations.lists.filter((l) => l.source.kind === 'unsupported').length;
      const parts: string[] = [];
      if (validations.lists.length) {
        parts.push(
          `${plural(validations.lists.length, 'list rule')} show a dropdown and check typed values as Excel does` +
            (unresolved ? ` (${unresolved} use a formula this editor can't resolve, so they have no dropdown)` : ''),
        );
      }
      if (validations.other) parts.push(`${plural(validations.other, 'other rule')} (numbers, dates, text length or custom) are kept in the file but not enforced here`);
      findings.push({
        kind: 'dataValidation',
        severity: 'info',
        title: 'Data validation',
        detail: `${parts.join('; ')}. Pasted values aren't checked, as in Excel.`,
        sheet: name,
      });
    }
    const linkCount = children(child(meta, 'hyperlinks'), 'hyperlink').length;
    if (linkCount) {
      findings.push({
        kind: 'hyperlink',
        severity: 'info',
        title: 'Hyperlinks',
        detail: `${plural(linkCount, 'link')}. Kept in the file; not clickable here.`,
        sheet: name,
      });
    }
    const sparkCount = countMatches(xml.slice(scan.sheetDataEnd), /<x14:sparklineGroup\b/g);
    if (sparkCount) {
      findings.push({
        kind: 'sparkline',
        severity: 'info',
        title: 'Sparklines',
        detail: `${plural(sparkCount, 'sparkline group')}. Kept in the file and redrawn by Excel; not shown here.`,
        sheet: name,
      });
    }
    if (child(meta, 'autoFilter')) {
      findings.push({
        kind: 'autoFilter',
        severity: 'info',
        title: 'Filter',
        detail: 'Kept in the file. Rows hidden by the filter stay hidden here, but the filter buttons are not shown.',
        sheet: name,
      });
    }
    if (scan.richTextCells) {
      findings.push({
        kind: 'richText',
        severity: 'info',
        title: 'Mixed formatting inside cells',
        detail: `${plural(scan.richTextCells, 'cell')} use more than one font style within their text. The text is shown plainly; editing one of those cells keeps its text but not the mixed formatting.`,
        sheet: name,
      });
    }

    for (const r of sheetRels) {
      if (r.type === 'drawing') {
        info.objects.push(...readDrawing(pkg, r.target));
      } else if (r.type === 'pivotTable') {
        const pxml = pkg.text(r.target);
        if (!pxml) continue;
        const pt = parseXml(pxml);
        const pname = attr(pt, 'name') ?? 'Pivot table';
        const ref = attr(child(pt, 'location'), 'ref') ?? '';
        const range = parseRange(ref);
        const cacheRel = pkg.rels(r.target).find((x) => x.type === 'pivotCacheDefinition');
        const source = cacheRel ? pivotSources.get(cacheRel.target) : undefined;
        if (range) {
          blocked.push({ ...range, kind: 'pivot', reason: `This cell is part of pivot table "${pname}". Change the pivot table in Excel.` });
          info.objects.push({ kind: 'pivot', name: pname, from: { row: range.r1, col: range.c1, rowOffPx: 0, colOffPx: 0 } });
          info.maxRow = Math.max(info.maxRow, range.r2);
          info.maxCol = Math.max(info.maxCol, range.c2);
        }
        findings.push({
          kind: 'pivot',
          severity: 'warn',
          title: `Pivot table "${pname}"`,
          detail: `Built from ${source ?? 'another range'}. Kept in the file and set to refresh when the file is opened in Excel, so it picks up your edits. Its cells can't be edited here.`,
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
        if (range && headerRows > 0) {
          blocked.push({
            r1: range.r1,
            c1: range.c1,
            r2: range.r1,
            c2: range.c2,
            kind: 'tableHeader',
            reason: `This is a column name in table "${tname}". Excel stores table column names separately, so rename the column in Excel.`,
          });
        }
        findings.push({
          kind: 'table',
          severity: 'info',
          title: `Table "${tname}"`,
          detail:
            'Kept in the file. Its column names (header row) can’t be edited here, and rows typed directly below it are not added to the table.',
          sheet: name,
          where: ref,
        });
      } else if (r.type === 'comments') {
        const n = countMatches(pkg.text(r.target), /<(\w+:)?comment\b/g);
        if (n) {
          findings.push({
            kind: 'comment',
            severity: 'info',
            title: 'Comments and notes',
            detail: `${plural(n, 'note')}. Kept in the file; not shown in this editor.`,
            sheet: name,
          });
        }
      } else if (r.type === 'ctrlProp') {
        info.objects.push({ kind: 'control', name: 'Form control' });
      } else if (r.type === 'slicer' || r.type === 'timeline') {
        findings.push({
          kind: 'slicer',
          severity: 'info',
          title: r.type === 'slicer' ? 'Slicers' : 'Timelines',
          detail: 'Kept in the file; not shown in this editor.',
          sheet: name,
        });
      }
    }

    const charts = info.objects.filter((o) => o.kind === 'chart');
    for (const c of charts) {
      const at = c.from ? toA1(c.from.row, c.from.col) : undefined;
      findings.push({
        kind: 'chart',
        severity: 'warn',
        title: c.title ? `Chart "${c.title}"` : 'Chart',
        detail: `${c.chartType ? c.chartType + ' chart' : 'Chart'}. Kept in the file and redrawn by Excel from the edited data when the file is opened. Shown here as a placeholder only.`,
        sheet: name,
        where: at,
      });
    }
    const images = info.objects.filter((o) => o.kind === 'image').length;
    if (images) {
      findings.push({
        kind: 'image',
        severity: 'info',
        title: 'Pictures',
        detail: `${plural(images, 'picture')}. Kept in the file; shown here as placeholders.`,
        sheet: name,
      });
    }
    const shapes = info.objects.filter((o) => o.kind === 'shape').length;
    if (shapes) {
      findings.push({
        kind: 'shape',
        severity: 'info',
        title: 'Shapes and text boxes',
        detail: `${plural(shapes, 'shape')}. Kept in the file; shown here as placeholders.`,
        sheet: name,
      });
    }
    const controls = info.objects.filter((o) => o.kind === 'control').length;
    if (controls) {
      findings.push({
        kind: 'control',
        severity: 'info',
        title: 'Form controls',
        detail: `${plural(controls, 'button or form control')}. Kept in the file; not shown or usable here.`,
        sheet: name,
      });
    }

    info.blocked = blocked;
    return { info, path, xml, scan, colStyles };
  });

  // Workbook-level features.
  const vba = relOfType('vbaProject').length > 0 || pkg.has('xl/vbaProject.bin');
  if (vba) {
    findings.push({
      kind: 'macro',
      severity: 'warn',
      title: 'Macros (VBA)',
      detail: 'This workbook contains macros. They are kept in the file but cannot run here.',
    });
  }
  const ext: Relationship[] = relOfType('externalLink');
  if (ext.length) {
    findings.push({
      kind: 'externalLink',
      severity: 'warn',
      title: 'Links to other workbooks',
      detail: `${plural(ext.length, 'linked workbook')}. Kept in the file. Formulas that use them show Excel's last value and are not recalculated here.`,
    });
  }
  if (relOfType('connections').length || pkg.list().some((p) => p.startsWith('xl/queryTables/'))) {
    findings.push({
      kind: 'connection',
      severity: 'info',
      title: 'Data connections',
      detail: 'External data connections or Power Query. Kept in the file; not refreshed here.',
    });
  }
  if (relOfType('slicerCache').length || relOfType('timelineCache').length) {
    if (!findings.some((f) => f.kind === 'slicer')) {
      findings.push({ kind: 'slicer', severity: 'info', title: 'Slicers', detail: 'Kept in the file; not shown in this editor.' });
    }
  }
  if (pkg.list().some((p) => p.startsWith('xl/richData/'))) {
    findings.push({
      kind: 'cellImage',
      severity: 'info',
      title: 'Pictures or data types in cells',
      detail: 'Some cells hold pictures or linked data types. Kept in the file; shown as their fallback value here.',
    });
  }
  if (child(wb, 'workbookProtection')) {
    findings.push({
      kind: 'protection',
      severity: 'info',
      title: 'Workbook structure is protected',
      detail: 'Kept in the file. This editor does not add, remove or rename sheets anyway.',
    });
  }

  return { workbookPath, sheets, styles, names, date1904, findings };
}

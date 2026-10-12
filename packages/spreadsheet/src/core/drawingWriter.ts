/**
 * Writes a sheet's charts, pictures and shapes as they now are into its
 * drawing part: objects deleted taken out, those moved or resized given
 * their new anchors, and charts added in the editor written as chart parts
 * of their own. Anchors the editor doesn't know are left as they are.
 */
import { anchorXml, chartAnchorXml } from './chartWriter';
import { addRel, removePart, setOverride } from './commentWriter';
import { placeSheetElement, rootTag, siblings } from './hyperlinks';
import type { SheetObject } from './types';
import { firstElement } from './xml';
import type { OoxmlPackage } from './zip';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_REL = `${REL}/drawing`;
const CHART_REL = `${REL}/chart`;
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The first free path like xl/charts/chart7.xml. */
function freePath(pkg: OoxmlPackage, dir: string, base: string): string {
  let n = 1;
  while (pkg.has(`${dir}/${base}${n}.xml`)) n++;
  return `${dir}/${base}${n}.xml`;
}

const sameAnchor = (a: SheetObject['from'], b: SheetObject['from']) =>
  a?.row === b?.row && a?.col === b?.col && a?.rowOffPx === b?.rowOffPx && a?.colOffPx === b?.colOffPx;

/** Takes a chart's part away, with the parts it brings (its style and colours). */
function removeChart(pkg: OoxmlPackage, drawing: string, anchor: string): void {
  // The chart element's relationship id, whatever its prefixes.
  const id = /<(?:\w+:)?chart\b[^>]*?\s\w+:id="([^"]+)"/.exec(anchor)?.[1];
  const rel = id ? pkg.rels(drawing).find((r) => r.id === id) : undefined;
  if (!rel || rel.type !== 'chart') return;
  for (const own of pkg.rels(rel.target)) if (!own.external) removePart(pkg, rel.target, own);
  removePart(pkg, drawing, rel);
}

/**
 * The sheet XML once its objects are written: `file` as the file had them,
 * `now` as they are. A sheet without a drawing part gets one for its first
 * new chart.
 */
export function writeObjects(pkg: OoxmlPackage, sheetPath: string, sheetXml: string, prefix: string, file: SheetObject[], now: SheetObject[]): string {
  const added = now.filter((o) => o.chartXml);
  const rel = pkg.rels(sheetPath).find((r) => r.type === 'drawing');
  let xml = sheetXml;
  let drawing = rel?.target;
  let dxml = drawing ? pkg.text(drawing) : undefined;
  if (!drawing || dxml === undefined) {
    if (!added.length) return xml;
    drawing = freePath(pkg, 'xl/drawings', 'drawing');
    dxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}"></xdr:wsDr>`;
    setOverride(pkg, drawing, DRAWING_TYPE);
    const id = addRel(pkg, sheetPath, DRAWING_REL, drawing);
    xml = placeSheetElement(xml, 'drawing', `<${prefix}drawing xmlns:r="${REL}" r:id="${id}"/>`);
  }
  const root = rootTag(dxml);
  if (!root) return xml;
  const p = root.qname.includes(':') ? root.qname.slice(0, root.qname.indexOf(':') + 1) : '';
  const kids = siblings(dxml, root.end + 1);
  const kept = new Map(now.filter((o) => o.source?.drawing === drawing).map((o) => [o.source!.anchor, o]));

  // From the last anchor back, so the earlier ones stay where they are.
  let out = dxml;
  for (const was of [...file].filter((o) => o.source?.drawing === drawing).sort((a, b) => b.source!.anchor - a.source!.anchor)) {
    const span = kids[was.source!.anchor];
    if (!span) continue;
    const anchor = out.slice(span.start, span.end);
    const is = kept.get(was.source!.anchor);
    if (!is) {
      removeChart(pkg, drawing, anchor);
      out = out.slice(0, span.start) + out.slice(span.end);
    } else if (!sameAnchor(was.from, is.from) || !sameAnchor(was.to, is.to)) {
      let moved = anchor;
      for (const [tag, at] of [['from', is.from], ['to', is.to]] as const) {
        const el = at && firstElement(moved, tag);
        if (el && at) moved = moved.slice(0, el.start) + anchorXml(p, tag, at) + moved.slice(el.end);
      }
      out = out.slice(0, span.start) + moved + out.slice(span.end);
    }
  }

  if (added.length) {
    let id = Math.max(1, ...[...out.matchAll(/cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])));
    let anchors = '';
    for (const o of added) {
      const part = freePath(pkg, 'xl/charts', 'chart');
      pkg.setText(part, o.chartXml!);
      setOverride(pkg, part, CHART_TYPE);
      const r = addRel(pkg, drawing, CHART_REL, part);
      anchors += chartAnchorXml(p, o.from!, o.to!, ++id, o.name, r);
    }
    const close = out.lastIndexOf('</');
    out = root.empty ? out.replace(/\/>\s*$/, `>${anchors}</${root.qname}>`) : out.slice(0, close) + anchors + out.slice(close);
  }
  if (out !== dxml) pkg.setText(drawing, out);
  return xml;
}

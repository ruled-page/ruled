import { attr, children, escapeXml, parseXml } from './xml';
import type { OoxmlPackage } from './zip';

const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
/** The format id every custom document property uses. */
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const DEFAULT_PATH = 'docProps/custom.xml';

function customPath(pkg: OoxmlPackage): string | undefined {
  return pkg.rels('').find((r) => r.type === 'custom-properties')?.target;
}

/** The workbook's custom document properties (File › Properties › Custom in Excel), as text. */
export function readCustomProperties(pkg: OoxmlPackage): Record<string, string> {
  const path = customPath(pkg);
  const xml = path ? pkg.text(path) : undefined;
  const out: Record<string, string> = {};
  if (!xml) return out;
  for (const p of children(parseXml(xml), 'property')) {
    const name = attr(p, 'name');
    const value = p.children[0];
    if (name && value) out[name] = value.text;
  }
  return out;
}

/**
 * Sets custom document properties as text, adding the part (and its
 * relationship and content type) when the workbook has none. Other
 * properties are left as they are.
 */
export function setCustomProperties(pkg: OoxmlPackage, props: Record<string, string>): void {
  const entries = Object.entries(props);
  if (!entries.length) return;
  let path = customPath(pkg);
  let xml = path ? pkg.text(path) : undefined;
  if (!path || !xml) {
    path = DEFAULT_PATH;
    xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>';
    const rels = pkg.text('_rels/.rels');
    if (rels) {
      pkg.setText('_rels/.rels', rels.replace(/<\/(\w+:)?Relationships>/, (close) => `<Relationship Id="rIdWbeCustom" Type="${REL_TYPE}" Target="${DEFAULT_PATH}"/>${close}`));
    }
    const ct = pkg.text('[Content_Types].xml');
    if (ct && !ct.includes(`PartName="/${DEFAULT_PATH}"`)) {
      pkg.setText('[Content_Types].xml', ct.replace(/<\/(\w+:)?Types>/, (close) => `<Override PartName="/${DEFAULT_PATH}" ContentType="${CONTENT_TYPE}"/>${close}`));
    }
  }
  const vt = /xmlns:(\w+)="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/docPropsVTypes"/.exec(xml)?.[1] ?? 'vt';
  const pids = [...xml.matchAll(/\bpid="(\d+)"/g)].map((m) => Number(m[1]));
  let pid = Math.max(1, ...pids);
  for (const [name, value] of entries) {
    const el = (id: number) => `<property fmtid="${FMTID}" pid="${id}" name="${escapeXml(name)}"><${vt}:lpwstr>${escapeXml(value)}</${vt}:lpwstr></property>`;
    const existing = new RegExp(`<(\\w+:)?property\\b[^>]*\\bname="${escapeXml(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>[\\s\\S]*?</(\\w+:)?property>`);
    const m = existing.exec(xml);
    if (m) {
      const id = Number(/\bpid="(\d+)"/.exec(m[0])?.[1] ?? ++pid);
      xml = xml.replace(m[0], el(id));
    } else {
      xml = xml.replace(/<\/(\w+:)?Properties>/, (close) => `${el(++pid)}${close}`);
    }
  }
  if (!xml.includes('xmlns:vt=') && vt === 'vt') xml = xml.replace(/<((\w+:)?Properties)\b/, '<$1 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"');
  pkg.setText(path, xml);
}

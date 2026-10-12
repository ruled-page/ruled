import { attr, child, children, descendants, elements, firstElement, parseAttrs, parseXml, type XmlNode } from './xml';

/*
 * Excel 365's checkboxes in cells. A checkbox is a cell holding TRUE or
 * FALSE whose format carries an extension, an xfComplement, pointing through
 * the workbook's feature property bags at a Checkbox cell control:
 *
 *   XFComplements bag → its Nth mapped bag, an XFComplement
 *     → its XFControls bag → its CellControl bag, of type Checkbox.
 */

/** The extension on a format (<xf>) that gives it a feature property bag. */
export const CHECKBOX_XF_EXT = '{C7286773-470A-42A8-94C5-96B5CB345126}';
export const BAGS_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';
export const BAGS_REL = 'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag';
export const BAGS_CONTENT_TYPE = 'application/vnd.ms-excel.featurepropertybag+xml';
/** Where Excel puts the bags, for a workbook that has none yet. */
export const BAGS_PART = 'xl/featurePropertyBag/featurePropertyBag.xml';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** The bag a bag refers to by key, as <bagId k="key">n</bagId>. */
function bagRef(bags: XmlNode[], bag: XmlNode | undefined, key: string): XmlNode | undefined {
  const id = children(bag, 'bagId').find((b) => attr(b, 'k') === key);
  return id ? bags[Number(id.text)] : undefined;
}

/** The bag ids the XFComplements bag maps xfComplement indexes to, in order. */
function mappedBags(bags: XmlNode[]): number[] {
  const complements = bags.find((b) => attr(b, 'type') === 'XFComplements');
  const list = children(complements, 'a').find((a) => attr(a, 'k') === 'MappedFeaturePropertyBags');
  return children(list, 'bagId').map((b) => Number(b.text));
}

/** The xfComplement index that makes a format a checkbox, if the workbook's bags have one. */
export function checkboxComplement(xml: string | undefined): number | undefined {
  if (!xml) return undefined;
  let root: XmlNode;
  try {
    root = parseXml(xml);
  } catch {
    return undefined;
  }
  const bags = children(root, 'bag');
  const i = mappedBags(bags).findIndex((id) => {
    const complement = bags[id];
    if (attr(complement, 'type') !== 'XFComplement') return false;
    const controls = bagRef(bags, complement, 'XFControls');
    return attr(controls, 'type') === 'XFControls' && attr(bagRef(bags, controls, 'CellControl'), 'type') === 'Checkbox';
  });
  return i < 0 ? undefined : i;
}

/** A format's xfComplement index, if it has one. */
export function xfComplement(xf: XmlNode): number | undefined {
  const ext = children(child(xf, 'extLst'), 'ext').find((e) => attr(e, 'uri') === CHECKBOX_XF_EXT);
  const i = attr(descendants(ext, 'xfComplement')[0], 'i');
  return i === undefined ? undefined : Number(i);
}

/**
 * The bags with a checkbox's added when they have none, as Excel writes
 * them, and the xfComplement index that points at it.
 */
export function withCheckboxBags(xml: string | undefined): { xml: string; index: number } {
  const known = checkboxComplement(xml);
  if (xml && known !== undefined) return { xml, index: known };
  if (!xml) {
    return {
      xml:
        `${XML_HEAD}<FeaturePropertyBags xmlns="${BAGS_NS}">` +
        '<bag type="Checkbox"/>' +
        '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>' +
        '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>' +
        '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>' +
        '</FeaturePropertyBags>',
      index: 0,
    };
  }
  // Bags of other kinds already there: the checkbox's go after them, mapped last.
  const root = parseXml(xml);
  const p = root.qname.includes(':') ? root.qname.slice(0, root.qname.indexOf(':') + 1) : '';
  const bags = children(root, 'bag');
  const n = bags.length;
  const mapped = mappedBags(bags);
  const added =
    `<${p}bag type="Checkbox"/>` +
    `<${p}bag type="XFControls"><${p}bagId k="CellControl">${n}</${p}bagId></${p}bag>` +
    `<${p}bag type="XFComplement"><${p}bagId k="XFControls">${n + 1}</${p}bagId></${p}bag>`;
  const close = xml.lastIndexOf('</');
  let out = xml.slice(0, close) + added + xml.slice(close);
  const complements = bags.findIndex((b) => attr(b, 'type') === 'XFComplements');
  if (complements < 0) {
    const end = out.lastIndexOf('</');
    out =
      out.slice(0, end) +
      `<${p}bag type="XFComplements" extRef="XFComplementsMapperExtRef"><${p}a k="MappedFeaturePropertyBags"><${p}bagId>${n + 2}</${p}bagId></${p}a></${p}bag>` +
      out.slice(end);
    return { xml: out, index: 0 };
  }
  // The mapped list is the XFComplements bag's <a k="MappedFeaturePropertyBags">.
  const bag = elements(out, 'bag')[complements];
  const list = elements(out, 'a', bag.start, bag.end).find((a) => parseAttrs(a.attrs).k === 'MappedFeaturePropertyBags');
  if (!list) return { xml: out, index: mapped.length };
  const listClose = out.lastIndexOf('</', list.end - 1);
  out = out.slice(0, listClose) + `<${p}bagId>${n + 2}</${p}bagId>` + out.slice(listClose);
  return { xml: out, index: mapped.length };
}

/** A format's children without the checkbox extension (and without an <extLst> it leaves empty). */
export function withoutCheckboxExt(inner: string): string {
  const list = firstElement(inner, 'extLst');
  if (!list) return inner;
  const body = inner.slice(list.start, list.end);
  const kept = body.replace(/<(\w+:)?ext\b[^>]*\buri="\{C7286773-470A-42A8-94C5-96B5CB345126\}"[\s\S]*?<\/(\w+:)?ext>/, '');
  const empty = !/<(\w+:)?ext\b/.test(kept);
  return inner.slice(0, list.start) + (empty ? '' : kept) + inner.slice(list.end);
}

/** A format's children with the checkbox extension, pointing at xfComplement `index`; <extLst> comes last. */
export function withCheckboxExt(inner: string, index: number, p: string): string {
  const ext = `<${p}ext uri="${CHECKBOX_XF_EXT}" xmlns:xfpb="${BAGS_NS}"><xfpb:xfComplement i="${index}"/></${p}ext>`;
  const base = withoutCheckboxExt(inner);
  const list = firstElement(base, 'extLst');
  if (!list) return `${base}<${p}extLst>${ext}</${p}extLst>`;
  const close = base.lastIndexOf('</', list.end - 1);
  return base.slice(0, close) + ext + base.slice(close);
}

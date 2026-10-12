import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { checkboxComplement, withCheckboxBags, withCheckboxExt, withoutCheckboxExt } from '../src/core/checkboxes';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import type { Rect } from '../src/core/types';
import { checkboxState } from '../src/state/checkbox';
import type { WorkbookStore } from '../src/state/store';
import { strToU8, zipSync } from 'fflate';
import { open as openSample, parts } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Tasks!A1:A3 "Pack", "Ship", empty; B1 TRUE, B2 FALSE, B3 a number, B4 a formula giving TRUE. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Tasks');
  ws.getCell('A1').value = 'Pack';
  ws.getCell('A2').value = 'Ship';
  ws.getCell('B1').value = true;
  ws.getCell('B2').value = false;
  ws.getCell('B3').value = 7;
  ws.getCell('B4').value = { formula: 'LEN(A1)=4', result: true };
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const rect = (r1: number, c1: number, r2 = r1, c2 = c1): Rect => ({ r1, c1, r2, c2 });
const open = (b = bytes) => WorkbookSession.open(b, 'tasks.xlsx', 'en-US');
const style = (s: WorkbookSession, row: number, col: number) => s.model.styles.get(s.current(0, row, col).style);

/** The bags Excel writes for a workbook whose only feature is checkboxes. */
const EXCEL_BAGS =
  '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">' +
  '<bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag>' +
  '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>' +
  '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>' +
  '</FeaturePropertyBags>';

describe('the parts of a checkbox', () => {
  it('find the xfComplement that makes a format a checkbox, through the bags it maps', () => {
    expect(checkboxComplement(EXCEL_BAGS)).toBe(0);
    expect(checkboxComplement(undefined)).toBeUndefined();
    expect(checkboxComplement('<FeaturePropertyBags xmlns="x"><bag type="Other"/></FeaturePropertyBags>')).toBeUndefined();
    expect(checkboxComplement('not xml <')).toBeUndefined();
  });

  it("are made as Excel makes them, or added after other bags, mapped last", () => {
    const made = withCheckboxBags(undefined);
    expect(made.index).toBe(0);
    expect(checkboxComplement(made.xml)).toBe(0);
    expect(withCheckboxBags(EXCEL_BAGS)).toEqual({ xml: EXCEL_BAGS, index: 0 });
    // Bags for another feature, with their own complement mapped first.
    const other =
      '<FeaturePropertyBags xmlns="x"><bag type="Other"/><bag type="XFComplement"><bagId k="Other">0</bagId></bag>' +
      '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>1</bagId></a></bag></FeaturePropertyBags>';
    const added = withCheckboxBags(other);
    expect(added.index).toBe(1);
    expect(checkboxComplement(added.xml)).toBe(1);
    const unmapped = withCheckboxBags('<FeaturePropertyBags xmlns="x"><bag type="Other"/></FeaturePropertyBags>');
    expect([unmapped.index, checkboxComplement(unmapped.xml)]).toEqual([0, 0]);
  });

  it("put the format's extension last among its children, beside any others, and take it off again", () => {
    const on = withCheckboxExt('<alignment horizontal="center"/>', 0, '');
    expect(on).toBe(
      '<alignment horizontal="center"/><extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><xfpb:xfComplement i="0"/></ext></extLst>',
    );
    expect(withoutCheckboxExt(on)).toBe('<alignment horizontal="center"/>');
    const other = '<extLst><ext uri="{other}"><x:y/></ext></extLst>';
    const both = withCheckboxExt(other, 2, '');
    expect(both).toMatch(/^<extLst><ext uri="\{other\}"><x:y\/><\/ext><ext uri="\{C7286773[^"]*"[^>]*><xfpb:xfComplement i="2"\/><\/ext><\/extLst>$/);
    expect(withoutCheckboxExt(both)).toBe(other);
  });

  it('show ticked for TRUE, unticked for FALSE or nothing, and not at all for other values', () => {
    const box = { checkbox: true } as never;
    expect(checkboxState({ t: 3, n: 1, style: 0, flags: 0 }, box)).toBe(true);
    expect(checkboxState({ t: 3, n: 0, style: 0, flags: 0 }, box)).toBe(false);
    expect(checkboxState(null, box)).toBe(false);
    expect(checkboxState({ t: 1, n: 7, style: 0, flags: 0 }, box)).toBeUndefined();
    expect(checkboxState({ t: 3, n: 1, style: 0, flags: 0 }, {} as never)).toBeUndefined();
  });
});

describe('checkboxes in a workbook', () => {
  it('are saved as Excel saves them, read back, reused, and taken off', () => {
    const s = open();
    s.apply({ type: 'format', sheet: 0, rects: [rect(0, 1, 1, 1)], change: { checkbox: true } });
    const saved = s.save();
    const p = parts(saved);
    expect(p['xl/featurePropertyBag/featurePropertyBag.xml']).toContain('<bag type="Checkbox"/>');
    expect(p['xl/_rels/workbook.xml.rels']).toMatch(/Type="http:\/\/schemas\.microsoft\.com\/office\/2022\/11\/relationships\/FeaturePropertyBag" Target="featurePropertyBag\/featurePropertyBag\.xml"/);
    expect(p['[Content_Types].xml']).toContain('PartName="/xl/featurePropertyBag/featurePropertyBag.xml" ContentType="application/vnd.ms-excel.featurepropertybag+xml"');
    expect(p['xl/styles.xml']).toContain('<xfpb:xfComplement i="0"/>');

    const t = open(saved);
    expect([style(t, 0, 1).checkbox, style(t, 1, 1).checkbox, style(t, 2, 1).checkbox]).toEqual([true, true, undefined]);
    // A second checkbox uses the bags already there.
    t.apply({ type: 'format', sheet: 0, rects: [rect(2, 1)], change: { checkbox: true } });
    const again = parts(t.save());
    expect(again['xl/featurePropertyBag/featurePropertyBag.xml']).toBe(p['xl/featurePropertyBag/featurePropertyBag.xml']);
    expect(again['xl/_rels/workbook.xml.rels'].match(/FeaturePropertyBag"/g)).toHaveLength(1);
    // Taken off, the values stay.
    t.apply({ type: 'format', sheet: 0, rects: [rect(0, 1, 2, 1)], change: { checkbox: false } });
    const u = open(t.save());
    expect(style(u, 0, 1).checkbox).toBeUndefined();
    expect(u.current(0, 0, 1).content).toEqual({ kind: 'bool', value: true });
  });
});

describe("a workbook XlsxWriter made, as Excel 365 does", () => {
  it('has its checkboxes read: ticked, unticked, and one a formula works out', () => {
    const s = openSample('kitchen-sink.xlsx');
    const at = (ref: string) => {
      const col = ref.charCodeAt(0) - 65;
      const row = Number(ref.slice(1)) - 1;
      return { style: s.model.styles.get(s.current(0, row, col).style), content: s.current(0, row, col).content };
    };
    expect(['H5', 'H6', 'H7', 'H8', 'H9'].map((r) => at(r).style.checkbox)).toEqual([true, true, true, true, true]);
    expect(at('H5').content).toEqual({ kind: 'bool', value: true });
    expect(at('H9').content).toMatchObject({ kind: 'formula' });
    expect(at('H4').style.checkbox).toBeUndefined();
  });
});

describe('a workbook with bags for something else', () => {
  it('gets the checkbox added to its own bags part, mapped after the others', () => {
    const made = open();
    made.apply({ type: 'format', sheet: 0, rects: [rect(0, 1)], change: { checkbox: true } });
    const p = parts(made.save());
    // The same package, its bags only another feature's: B1's complement points at that.
    p['xl/featurePropertyBag/featurePropertyBag.xml'] =
      '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><bag type="Other"/>' +
      '<bag type="XFComplement"><bagId k="Other">0</bagId></bag>' +
      '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>1</bagId></a></bag></FeaturePropertyBags>';
    const s = open(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))));
    expect(style(s, 0, 1).checkbox).toBeUndefined();
    s.apply({ type: 'format', sheet: 0, rects: [rect(1, 1)], change: { checkbox: true } });
    const saved = s.save();
    const again = parts(saved);
    expect(again['xl/featurePropertyBag/featurePropertyBag.xml']).toContain('<bag type="Checkbox"/>');
    expect(again['xl/_rels/workbook.xml.rels'].match(/FeaturePropertyBag"/g)).toHaveLength(1);
    const t = open(saved);
    expect([style(t, 0, 1).checkbox, style(t, 1, 1).checkbox]).toEqual([undefined, true]);
  });
});

describe('checkboxes in the editor', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());
  const select = (r: Rect) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
  const ticks = () => [0, 1, 2, 3].map((row) => store.checkboxAt(row, 1));

  it('are inserted with empty cells unticked, ticked by click or Space, and Delete unticks then removes them, each a step', async () => {
    store = await testStore();
    await opened(store, bytes, 'tasks.xlsx');
    select(rect(0, 1, 4, 1));
    await store.insertCheckboxes();
    // A number shows as itself; the empty B5 is now FALSE.
    expect([...ticks(), store.checkboxAt(4, 1)]).toEqual([true, false, undefined, true, false]);
    expect(store.view!.get(4, 1)?.t).toBe(3);
    expect(store.activeStyle.checkbox).toBe(true);

    // Space sets the selection the other way from the active cell, leaving formulas alone.
    await store.toggleCheckboxes();
    expect([...ticks(), store.checkboxAt(4, 1)]).toEqual([false, false, undefined, true, false]);
    await store.toggleCheckboxes({ row: 1, col: 1 });
    expect(store.checkboxAt(1, 1)).toBe(true);
    expect(await store.toggleCheckboxes({ row: 2, col: 1 })).toBe(false);

    // Delete unticks first, then takes the checkboxes and their values away.
    await store.clear([rect(0, 1, 4, 1)]);
    expect([...ticks(), store.checkboxAt(4, 1)]).toEqual([false, false, undefined, true, false]);
    await store.clear([rect(0, 1, 2, 1)]);
    expect(ticks()).toEqual([undefined, undefined, undefined, true]);
    expect(store.view!.get(0, 1)?.t ?? 0).toBe(0);
    await store.undo();
    expect(ticks()).toEqual([false, false, undefined, true]);
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Insert checkboxes', 'Untick 3 checkboxes', 'Tick a checkbox', 'Untick a checkbox', 'Remove checkboxes', 'Undo']);
  });
});

import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { VT } from '../src/core/types';
import { english } from '../src/i18n';
import type { Answer, WorkbookStore } from '../src/state/store';
import {
  defineName,
  deleteName,
  listedNames,
  nameBoxEnter,
  nameBoxNames,
  nameError,
  nameHasError,
  nameValue,
  refersToError,
  refersToFile,
  selectionName,
  selectionRefersTo,
} from '../src/state/names';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * Data!A1:A3 = 10, 20, 30, B1 = "North", B2 = "South"; Calc!A1 = SUM(Values),
 * A2 = Here, C1 = 5. Names: Values (Data!A1:A3), Labels (Data!B1:B2), Rate
 * (0.2, a constant), Dynamic (an OFFSET), Broken (#REF!), Here (the
 * workbook's, Data!A1, and Calc's own, Calc!C1), Pair (two areas), Secret
 * (hidden) and a print area.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  [10, 20, 30].forEach((v, i) => (data.getCell(`A${i + 1}`).value = v));
  data.getCell('B1').value = 'North';
  data.getCell('B2').value = 'South';
  const calc = wb.addWorksheet('Calc');
  calc.getCell('A1').value = { formula: 'SUM(Values)', result: 60 };
  calc.getCell('A2').value = { formula: 'Here', result: 5 };
  calc.getCell('C1').value = 5;
  const files = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  files['xl/workbook.xml'] = strToU8(
    strFromU8(files['xl/workbook.xml']).replace(
      '</sheets>',
      '</sheets><definedNames>' +
        '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$3</definedName>' +
        '<definedName name="Values">Data!$A$1:$A$3</definedName>' +
        '<definedName name="Labels" comment="Regions">Data!$B$1:$B$2</definedName>' +
        '<definedName name="Rate">0.2</definedName>' +
        '<definedName name="Dynamic">OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)</definedName>' +
        '<definedName name="Broken">#REF!</definedName>' +
        '<definedName name="here">Data!$A$1</definedName>' +
        '<definedName name="Here" localSheetId="1">Calc!$C$1</definedName>' +
        '<definedName name="Pair">Data!$A$1,Data!$B$2</definedName>' +
        '<definedName name="Secret" hidden="1">Data!$A$2</definedName>' +
        '</definedNames>',
    ),
  );
  return zipSync(files);
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'names.xlsx');
  return store;
}

const DATA = 0;
const CALC = 1;
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (...rects: ReturnType<typeof rect>[]) => store.setSelection({ row: rects[0].r1, col: rects[0].c1, rects });
const toasts = () => store.toasts.map((t) => english.text(t.text));
const actions = () => store.history.map((h) => english.text(h.action));
const value = (sheet: number, row: number, col: number) => {
  const rec = store.sheets[sheet].get(row, col);
  return rec?.t === VT.Error ? rec.s : rec?.n;
};
const where = () => ({ sheet: store.active, rects: store.selection.rects });
const text = (t: unknown) => (t === undefined ? undefined : english.text(t as never));

async function answer(reply: Answer): Promise<void> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  store.confirmRequest!.resolve(reply);
}

async function saved() {
  const s = await store.snapshot();
  return WorkbookSession.open(s!.bytes, 'names.xlsx').model.names;
}

describe('the names a workbook has', () => {
  it('arrive with it, and are listed as Excel lists them: alphabetically, without hidden names or Excel’s own', async () => {
    await open();
    expect(store.names).toHaveLength(10);
    expect(listedNames(store).map((n) => `${n.name}${n.localSheet === undefined ? '' : `@${n.localSheet}`}`)).toEqual([
      'Broken',
      'Dynamic',
      'here',
      'Here@1',
      'Labels',
      'Pair',
      'Rate',
      'Values',
    ]);
  });

  it('that refer to cells are offered by the name box, a sheet’s own name hiding the workbook’s', async () => {
    await open();
    expect(nameBoxNames(store)).toEqual(['here', 'Labels', 'Pair', 'Values']);
    store.setActive(CALC);
    expect(nameBoxNames(store)).toEqual(['Here', 'Labels', 'Pair', 'Values']);
  });

  it('are previewed as Excel’s Name Manager does, and those with errors are known', async () => {
    await open();
    const preview = (name: string, localSheet?: number) => nameValue(store, store.names.find((n) => n.name === name && n.localSheet === localSheet)!);
    expect(preview('Values')).toBe('{10;20;30}');
    expect(preview('Labels')).toBe('{"North";"South"}');
    expect(preview('here')).toBe('10');
    expect(preview('Here', CALC)).toBe('5');
    expect(preview('Rate')).toBe('0.2');
    expect(preview('Dynamic')).toBe('');
    expect(preview('Broken')).toBe('#REF!');
    expect(preview('Pair')).toBe('');
    expect(nameValue(store, { name: 'Text', formula: '"hi"' })).toBe('"hi"');
    expect(nameValue(store, { name: 'Many', formula: 'Data!$A:$A' })).toBe('{10;20;30…}');
    expect(nameValue(store, { name: 'Wide', formula: 'Data!$A$1:$E$5' })).toBe('{10,"North";20,"South";30,…}');
    expect(listedNames(store).filter((n) => nameHasError(store, n)).map((n) => n.name)).toEqual(['Broken']);
  });
});

describe('going to a name', () => {
  it('selects the range it refers to, switching sheet, whatever the case it is typed in', async () => {
    await open();
    store.setActive(CALC);
    expect(store.goTo('values')).toBe(true);
    expect(where()).toEqual({ sheet: DATA, rects: [rect(0, 0, 2, 0)] });
    // On Calc, Here is Calc's own; elsewhere the workbook's, unless the sheet is named.
    store.setActive(CALC);
    expect(store.goTo('Here')).toBe(true);
    expect(where()).toEqual({ sheet: CALC, rects: [rect(0, 2)] });
    expect(store.goTo('HERE')).toBe(true);
    expect(where()).toEqual({ sheet: CALC, rects: [rect(0, 2)] });
    store.setActive(DATA);
    expect(store.goTo('Here')).toBe(true);
    expect(where()).toEqual({ sheet: DATA, rects: [rect(0, 0)] });
    expect(store.goTo('Calc!Here')).toBe(true);
    expect(where()).toEqual({ sheet: CALC, rects: [rect(0, 2)] });
    expect(store.goTo('Pair')).toBe(true);
    expect(where()).toEqual({ sheet: DATA, rects: [rect(0, 0), rect(1, 1)] });
    // Cell references still win, and hidden names aren't gone to.
    expect(store.goTo('B2')).toBe(true);
    expect(where()).toEqual({ sheet: DATA, rects: [rect(1, 1)] });
    expect(store.goTo('Secret')).toBe(false);
  });

  it('is refused, saying why, for a name that isn’t a range or text that is neither a reference nor a name', async () => {
    await open();
    expect(await nameBoxEnter(store, 'Rate')).toBe(false);
    expect(await nameBoxEnter(store, 'Dynamic')).toBe(false);
    expect(await nameBoxEnter(store, 'no such thing')).toBe(false);
    expect(await nameBoxEnter(store, 'Nowhere!A1')).toBe(false);
    expect(toasts()).toEqual([
      "Rate refers to =0.2, which isn't a range of cells, so it can't be selected.",
      "Dynamic refers to =OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1), which isn't a range of cells, so it can't be selected.",
      '"no such thing" isn\'t a cell, a range or a name in this workbook.',
      '"Nowhere!A1" isn\'t a cell, a range or a name in this workbook.',
    ]);
    expect(store.names).toHaveLength(10);
  });

  it('shows in the name box when the selection is exactly a named range', async () => {
    await open();
    select(rect(0, 0, 2, 0));
    expect(selectionName(store)).toBe('Values');
    select(rect(0, 0));
    expect(selectionName(store)).toBe('here');
    select(rect(0, 0, 1, 0));
    expect(selectionName(store)).toBeUndefined();
    store.setActive(CALC);
    select(rect(0, 2));
    expect(selectionName(store)).toBe('Here');
  });
});

describe('naming the selection from the name box', () => {
  it('defines a name for the workbook, as one step to undo, listed in the changes and saved', async () => {
    await open();
    select(rect(0, 1, 1, 1));
    expect(selectionRefersTo(store)).toBe('=Data!$B$1:$B$2');
    store.setActive(CALC);
    select(rect(4, 3, 5, 4));
    expect(await nameBoxEnter(store, ' Totals ')).toBe(true);
    expect(store.names.find((n) => n.name === 'Totals')).toEqual({ name: 'Totals', formula: 'Calc!$D$5:$E$6' });
    expect(selectionName(store)).toBe('Totals');
    expect(actions()).toEqual(['Define name Totals']);
    expect(store.editedCount).toBe(1);
    await store.refreshChanges();
    expect(store.changeList.map((c) => english.text(c.note!))).toEqual(['Defined the name Totals as =Calc!$D$5:$E$6']);
    expect((await saved()).some((n) => n.name === 'Totals')).toBe(true);

    await store.undo();
    expect(store.names.some((n) => n.name === 'Totals')).toBe(false);
    expect(store.editedCount).toBe(0);
    await store.redo();
    expect(store.names.some((n) => n.name === 'Totals')).toBe(true);
    expect(store.goTo('totals')).toBe(true);
  });

  it('refuses a name that is already taken, or a sheet that isn’t a worksheet', async () => {
    await open();
    expect(await nameBoxEnter(store, 'secret')).toBe(false);
    expect(toasts()).toEqual(['The workbook already has a name "Secret".']);
    store.sheets[DATA].info.kind = 'chartsheet';
    expect(await nameBoxEnter(store, 'Fresh')).toBe(false);
    expect(toasts()[1]).toBe('Only cells on a worksheet can be named.');
  });
});

describe('defining and editing names', () => {
  it('checks the name and its scope, and the formula, before sending them', async () => {
    await open();
    expect(text(nameError(store, 'A1', undefined))).toBe('"A1" is a cell reference, so it can\'t be a name.');
    expect(text(nameError(store, 'VALUES', undefined))).toBe('The workbook already has a name "Values".');
    expect(text(nameError(store, 'Values', DATA))).toBeUndefined();
    expect(text(nameError(store, 'here', CALC))).toBe('Sheet Calc already has a name "Here".');
    expect(text(nameError(store, 'HERE', CALC, { name: 'Here', localSheet: CALC }))).toBeUndefined();
    expect(text(refersToError(store, '=Nowhere!A1'))).toBe('This workbook has no sheet called "Nowhere".');
    expect(text(refersToError(store, '=Calc!$A$1:$A$2'))).toBeUndefined();
    store.setActive(CALC);
    expect(refersToFile(store, '=$A$1+Rate')).toBe('Calc!$A$1+Rate');
    expect(text(await defineName(store, { name: 'Values', formula: '1' }))).toBe('The workbook already has a name "Values".');
    expect(store.history).toEqual([]);
  });

  it('local to a sheet, and edits one, renaming it in the formulas that use it', async () => {
    await open();
    expect(await defineName(store, { name: 'Values', formula: 'Data!$A$2:$A$3', localSheet: CALC, comment: 'Just two' })).toBeUndefined();
    expect(value(CALC, 0, 0)).toBe(50);
    expect(await defineName(store, { name: 'Amounts', formula: 'Data!$A$2:$A$3', localSheet: CALC, comment: 'Just two' }, { name: 'Values', localSheet: CALC })).toBeUndefined();
    expect(store.sheets[CALC].get(0, 0)?.f).toBe('SUM(Amounts)');
    expect(value(CALC, 0, 0)).toBe(50);
    expect(actions()).toEqual(['Define name Values', 'Edit name Amounts']);
    await store.undo();
    expect(store.sheets[CALC].get(0, 0)?.f).toBe('SUM(Values)');
    expect(value(CALC, 0, 0)).toBe(50);
    await store.undo();
    expect(value(CALC, 0, 0)).toBe(60);
  });
});

describe('deleting a name', () => {
  it('asks first, then leaves #NAME? in the formulas that used it', async () => {
    await open();
    const values = store.names.find((n) => n.name === 'Values')!;
    const kept = deleteName(store, values);
    await answer('cancel');
    expect(await kept).toBe(false);
    expect(store.names).toContain(values);

    const gone = deleteName(store, values);
    await answer('confirm');
    expect(await gone).toBe(true);
    expect(english.text(store.history[0].action)).toBe('Delete name Values');
    expect(value(CALC, 0, 0)).toBe('#NAME?');
    expect((await saved()).some((n) => n.name === 'Values')).toBe(false);
    await store.undo();
    expect(value(CALC, 0, 0)).toBe(60);
  });
});

describe('names and structural changes', () => {
  it('follow rows inserted above them', async () => {
    await open();
    select(rect(0, 0));
    await store.insertRows('above');
    expect(store.names.find((n) => n.name === 'Values')?.formula).toBe('Data!$A$2:$A$4');
    expect(store.goTo('Values')).toBe(true);
    expect(where()).toEqual({ sheet: DATA, rects: [rect(1, 0, 3, 0)] });
    await store.undo();
    expect(store.names.find((n) => n.name === 'Values')?.formula).toBe('Data!$A$1:$A$3');
  });
});

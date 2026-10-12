import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import { CF_PRESETS, quickRule } from '../src/state/cfRules';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Data!A1:A6 holds 1..6 with a rule making values over 4 bold; Locked is protected. */
let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let r = 1; r <= 6; r++) ws.getCell(`A${r}`).value = r;
  ws.addConditionalFormatting({ ref: 'A1:A6', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['4'], priority: 1, style: { font: { bold: true } } }] });
  await wb.addWorksheet('Locked').protect('', {});
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

let store: WorkbookStore;
afterEach(() => store?.dispose());
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
const rules = () => store.view!.info.conditionalFormats;
const actions = () => store.history.map((h) => english.text(h.action));

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'fmt.xlsx');
  return store;
}

describe('conditional formatting in the editor', () => {
  it('adds a rule first in the order, and shows its format once the worker has worked it out', async () => {
    await open();
    select(rect(0, 0, 5, 0));
    const made = quickRule('lessThan', ['3'], CF_PRESETS[2].style, store.validationPlace());
    if (!('rule' in made)) throw new Error('no rule');
    await store.addConditionalFormat(made.rule);
    expect(rules().map((r) => [r.operator, r.priority])).toEqual([
      ['lessThan', 1],
      ['greaterThan', 2],
    ]);
    const formatAt = (row: number) => store.view!.cfStyles[store.view!.cfAt(row, 0)?.style ?? -1];
    await vi.waitFor(() => expect(formatAt(0)).toMatchObject({ fill: '#C6EFCE' }));
    expect(actions()).toEqual(['Conditional formatting']);
  });

  it('clears rules from the selected cells, the rest keeping them, or from the whole sheet', async () => {
    await open();
    select(rect(2, 0, 3, 0));
    await store.clearConditionalFormats('selection');
    expect(rules()[0].rects).toEqual([rect(0, 0, 1, 0), rect(4, 0, 5, 0)]);
    await store.clearConditionalFormats('sheet');
    expect(rules()).toEqual([]);
    expect(store.view!.cfAt(5, 0)).toBeUndefined();
    expect(actions()).toEqual(['Clear conditional formatting', 'Clear conditional formatting']);
    await store.undo();
    expect(rules()).toHaveLength(1);
  });

  it("opens Excel's Rules Manager on the sheet's rules, and its quick dialogs on the selection", async () => {
    await open();
    store.openCfManager();
    expect(store.cfManager?.rules).toEqual(rules());
    select(rect(1, 1));
    store.openCfDialog('top');
    expect(store.cfDialog).toEqual({ kind: 'top', rects: [rect(1, 1)] });
  });

  it('is refused on a protected sheet', async () => {
    await open();
    store.setActive(1);
    store.openCfManager();
    store.openCfDialog('equal');
    await store.clearConditionalFormats('sheet');
    expect([store.cfManager, store.cfDialog]).toEqual([undefined, undefined]);
    expect(store.toasts.map((t) => english.text(t.text))).toEqual(Array(3).fill('This sheet is protected in Excel, so it is read-only here.'));
  });
});

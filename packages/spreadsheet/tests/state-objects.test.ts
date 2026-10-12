import ExcelJS from 'exceljs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { Anchor } from '../src/core/types';
import { english } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { objectEditable, objectKey, shiftAnchor } from '../src/ui/objectGeometry';
import { parts } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

const at = (row: number, col: number, rowOffPx = 0, colOffPx = 0): Anchor => ({ row, col, rowOffPx, colOffPx });

/** A tiny PNG, for a picture anchored at one cell with a size of its own. */
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');

/** Data!A1:C4 to chart, column E hidden, and a 100 by 50 picture at H2; Locked is protected, with a picture too. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow([null, 'Q1', 'Q2']);
  ws.addRow(['North', 10, 12]);
  ws.addRow(['South', 7, 9]);
  ws.addRow(['East', 4, 5]);
  ws.getColumn(5).hidden = true;
  const image = wb.addImage({ buffer: PNG as unknown as ExcelJS.Buffer, extension: 'png' });
  ws.addImage(image, { tl: { col: 7, row: 1 }, ext: { width: 100, height: 50 } });
  const locked = wb.addWorksheet('Locked');
  locked.addImage(image, { tl: { col: 1, row: 1 }, ext: { width: 40, height: 40 } });
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

describe('objects on the sheet', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());
  const open = async () => {
    store = await testStore();
    await opened(store, await workbook(), 'objects.xlsx');
    return store;
  };
  const objects = () => store.view!.info.objects;
  const history = () => store.history.map((h) => english.text(h.action));

  it('move their anchors along the rows and columns, past hidden ones, never before A1', async () => {
    await open();
    const view = store.view!;
    // Columns are 64 pixels; E (index 4) is hidden.
    expect(shiftAnchor(view, at(0, 2, 0, 10), 60, 25)).toEqual(at(1, 3, 5, 6));
    expect(shiftAnchor(view, at(0, 3, 0, 10), 60, 0)).toEqual(at(0, 5, 0, 6));
    expect(shiftAnchor(view, at(0, 5, 0, 6), -10, 0)).toEqual(at(0, 3, 0, 60));
    expect(shiftAnchor(view, at(1, 1, 5, 5), -500, -500)).toEqual(at(0, 0, 0, 0));
  });

  it('move, resize and delete as one step each, keeping a picture anchored at one cell so', async () => {
    await open();
    store.setSelection({ row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 3, c2: 2 }] });
    await store.insertChart('column');
    const [picture, chart] = objects();
    expect(picture).toMatchObject({ kind: 'image', from: at(1, 7), sizePx: { w: 100, h: 50 } });
    expect(store.editableObjects().map(objectKey)).toEqual([objectKey(picture), objectKey(chart)]);

    // The chart dragged 20 pixels right and 30 down; then its bottom-right handle 40 pixels further out.
    await store.placeObject(objectKey(chart), { left: 20, top: 30, right: 20, bottom: 30 });
    const moved = objects()[1];
    expect([moved.from, moved.to]).toEqual([shiftAnchor(store.view!, chart.from!, 20, 30), shiftAnchor(store.view!, chart.to!, 20, 30)]);
    await store.placeObject(objectKey(chart), { left: 0, top: 0, right: 40, bottom: 40 });
    expect(objects()[1].from).toEqual(moved.from);
    expect(objects()[1].to).toEqual(shiftAnchor(store.view!, moved.to!, 40, 40));
    // The picture keeps its one anchor and changes its size.
    await store.placeObject(objectKey(picture), { left: -10, top: 0, right: 30, bottom: 10 });
    expect(objects()[0]).toMatchObject({ from: at(1, 6, 0, 54), sizePx: { w: 140, h: 60 } });
    // Nothing moved is no step.
    await store.placeObject(objectKey(picture), { left: 0, top: 0, right: 0, bottom: 0 });
    expect(history()).toEqual(['Insert chart', 'Move Chart 1', 'Resize Chart 1', `Resize ${picture.name}`]);

    const saved = parts((await store.snapshot())!.bytes);
    expect(saved['xl/drawings/drawing1.xml']).toContain('<xdr:ext cx="1333500" cy="571500"/>');

    store.selectObject(objectKey(chart));
    await store.deleteObject(objectKey(chart));
    expect(objects().map((o) => o.kind)).toEqual(['image']);
    expect(store.selectedObject).toBeUndefined();
    await store.undo();
    expect(objects().map((o) => o.kind)).toEqual(['image', 'chart']);
  });

  it('change a chart from the dialog as one step, and leave a chart alone when nothing changed', async () => {
    await open();
    store.setSelection({ row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 3, c2: 2 }] });
    await store.insertChart('column');
    const chart = objects()[1];
    await store.editChart(objectKey(chart), {});
    await store.editChart(objectKey(chart), { kind: 'line', title: 'By region', legend: 'top' });
    const changed = objects()[1];
    expect([changed.title, english.text(changed.chartType!), changed.chart?.groups[0].kind, changed.chart?.legend?.position]).toEqual(['By region', 'Line', 'line', 'top']);
    expect(history()).toEqual(['Insert chart', 'Change Chart 1']);
    await store.undo();
    expect(objects()[1].chart?.groups[0].kind).toBe('bar');
  });

  it("can't be moved or deleted on a protected sheet", async () => {
    await open();
    store.setActive(1);
    const [picture] = objects();
    expect(objectEditable(store.view!.info, picture)).toBe(false);
    expect(store.editableObjects()).toEqual([]);
    await store.deleteObject(objectKey(picture));
    expect(objects()).toHaveLength(1);
  });
});

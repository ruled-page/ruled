import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { WorkerRequest, WorkerResponse } from '../src/worker/protocol';

/** Loads a fresh copy of the worker module and talks to it directly, collecting everything it posts. */
async function worker() {
  const posted: WorkerResponse[] = [];
  const scope = { onmessage: null as ((e: MessageEvent<WorkerRequest>) => void) | null, postMessage: (m: WorkerResponse) => posted.push(m) };
  (globalThis as { self?: unknown }).self = scope;
  vi.resetModules();
  await import('../src/worker/workbook.worker');
  let id = 0;
  /** Sends a request and returns what the worker answered to it. */
  const send = (req: Omit<WorkerRequest, 'id'>): WorkerResponse => {
    const message = { ...req, id: ++id } as WorkerRequest;
    scope.onmessage!({ data: message } as MessageEvent<WorkerRequest>);
    return posted.find((p) => 'id' in p && p.id === message.id)!;
  };
  const tick = () => new Promise((r) => setTimeout(r, 0));
  return { posted, send, tick };
}

let bytes: ArrayBuffer;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow([1, 2, { formula: 'A1+B1', result: 3 }]);
  bytes = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
});

const open = (send: Awaited<ReturnType<typeof worker>>['send']) => send({ type: 'open', bytes: bytes.slice(0), fileName: 'data.xlsx', locale: 'en-US' } as Omit<WorkerRequest, 'id'>);
const errorOf = (r: WorkerResponse) => (r.type === 'error' ? english.text(r.text ?? r.message) : undefined);

describe('the workbook worker', () => {
  it('turns down every request until a workbook is open', async () => {
    const { send } = await worker();
    const requests: Omit<WorkerRequest, 'id'>[] = [
      { type: 'edit', op: { type: 'input', sheet: 0, row: 0, col: 0, text: '1' } },
      { type: 'save' },
      { type: 'changes' },
      { type: 'validate', sheet: 0, row: 0, col: 0, text: '1', rule: 0 },
      { type: 'listSource', sheet: 0, formula: 'A1:A2' },
      { type: 'conditionalFormats', sheet: 0 },
      { type: 'structure', op: { type: 'insertRows', sheet: 0, at: 0, count: 1 } },
      { type: 'enableRecalc' },
    ] as Omit<WorkerRequest, 'id'>[];
    for (const req of requests) expect(errorOf(send(req)), req.type).toBe('No workbook is open.');
    expect(errorOf(send({ type: 'undoStructure' } as Omit<WorkerRequest, 'id'>))).toBe("That change can't be undone any more.");
  });

  it('says why a file could not be opened, in words to translate', async () => {
    const { send } = await worker();
    const r = send({ type: 'open', bytes: new TextEncoder().encode('hello').buffer, fileName: 'x.xlsx', locale: 'en-US' } as Omit<WorkerRequest, 'id'>);
    expect(r).toMatchObject({ type: 'error', text: { key: 'core.open.notWorkbook' } });
  });

  it('sends the engine report once, after the grid, and always before an edit result', async () => {
    const { send, posted, tick } = await worker();
    expect(open(send).type).toBe('opened');
    expect(posted.map((p) => p.type)).toEqual(['opened']);
    send({ type: 'edit', op: { type: 'input', sheet: 0, row: 0, col: 0, text: '5' } } as Omit<WorkerRequest, 'id'>);
    await tick();
    expect(posted.map((p) => p.type)).toEqual(['opened', 'calc', 'edited']);
  });

  it('builds the engine in the next task when nothing asks for it sooner', async () => {
    const { send, posted, tick } = await worker();
    open(send);
    await tick();
    expect(posted.map((p) => p.type)).toEqual(['opened', 'calc']);
  });

  it('answers lists, checks and conditional formats', async () => {
    const { send } = await worker();
    open(send);
    expect(send({ type: 'listSource', sheet: 0, formula: 'A1:B1' } as Omit<WorkerRequest, 'id'>)).toMatchObject({ type: 'listSource', source: { kind: 'values', values: ['1', '2'] } });
    expect(send({ type: 'listSource', sheet: 0, formula: 'INDIRECT("Nowhere")' } as Omit<WorkerRequest, 'id'>)).toMatchObject({ type: 'listSource', source: null });
    expect(send({ type: 'conditionalFormats', sheet: 0 } as Omit<WorkerRequest, 'id'>)).toMatchObject({ type: 'conditionalFormats', result: { cells: [] } });
    expect(send({ type: 'validate', sheet: 0, row: 0, col: 0, text: 'x', rule: 0 } as Omit<WorkerRequest, 'id'>)).toMatchObject({ type: 'validated', ok: true });
    expect(send({ type: 'changes' } as Omit<WorkerRequest, 'id'>)).toMatchObject({ type: 'changes', changes: [] });
  });

  it('goes back and forth between structural changes, sending the whole workbook each time', async () => {
    const { send, posted, tick } = await worker();
    open(send);
    await tick();
    const inserted = send({ type: 'structure', op: { type: 'insertRows', sheet: 0, at: 0, count: 1 } } as Omit<WorkerRequest, 'id'>);
    expect(inserted).toMatchObject({ type: 'restructured', editedCount: 1, report: undefined, warnings: [] });
    // Undone before the new workbook's engine was built, so that build is skipped; the old one has its engine.
    const undone = send({ type: 'undoStructure' } as Omit<WorkerRequest, 'id'>);
    expect(undone).toMatchObject({ type: 'restructured', editedCount: 0 });
    expect(undone.type === 'restructured' && undone.report).toBeTruthy();
    await tick();
    expect(posted.filter((p) => p.type === 'calc')).toHaveLength(1);

    const redone = send({ type: 'redoStructure' } as Omit<WorkerRequest, 'id'>);
    expect(redone).toMatchObject({ type: 'restructured', editedCount: 1 });
    expect(errorOf(send({ type: 'redoStructure' } as Omit<WorkerRequest, 'id'>))).toBe('There is nothing to redo.');
    await tick();
    expect(posted.filter((p) => p.type === 'calc')).toHaveLength(2);
  });

  it('drops what could be redone after a new structural change', async () => {
    const { send } = await worker();
    open(send);
    send({ type: 'structure', op: { type: 'insertRows', sheet: 0, at: 0, count: 1 } } as Omit<WorkerRequest, 'id'>);
    send({ type: 'undoStructure' } as Omit<WorkerRequest, 'id'>);
    send({ type: 'structure', op: { type: 'insertCols', sheet: 0, at: 0, count: 1 } } as Omit<WorkerRequest, 'id'>);
    expect(errorOf(send({ type: 'redoStructure' } as Omit<WorkerRequest, 'id'>))).toBe('There is nothing to redo.');
  });

  it('reports an edit it cannot make as an error, and carries on', async () => {
    const { send } = await worker();
    open(send);
    const r = send({ type: 'edit', op: { type: 'input', sheet: 7, row: 0, col: 0, text: '1' } } as Omit<WorkerRequest, 'id'>);
    expect(r.type).toBe('error');
    expect(send({ type: 'edit', op: { type: 'input', sheet: 0, row: 0, col: 0, text: '1' } } as Omit<WorkerRequest, 'id'>).type).toBe('edited');
  });

  it('turns live recalculation on', async () => {
    const { send } = await worker();
    send({ type: 'open', bytes: bytes.slice(0), fileName: 'data.xlsx', locale: 'en-US', options: { recalc: 'off' } } as Omit<WorkerRequest, 'id'>);
    const r = send({ type: 'enableRecalc' } as Omit<WorkerRequest, 'id'>);
    expect(r).toMatchObject({ type: 'recalcOn', report: { liveCells: 1 } });
    expect(r.type === 'recalcOn' && r.report.off).toBeUndefined();
  });
});

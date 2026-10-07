import { WorkbookSession } from '../core/session';
import type { WorkerRequest, WorkerResponse } from './protocol';

/** The workbook, its parsed sheets and the calculation engine live here, off the UI thread. */
let session: WorkbookSession | undefined;
let calcSent = false;

const scope = self as unknown as {
  postMessage(message: WorkerResponse): void;
  onmessage: ((e: MessageEvent<WorkerRequest>) => void) | null;
};

function post(message: WorkerResponse): void {
  scope.postMessage(message);
}

/** Builds the engine if needed and reports it, so its results always precede edit results. */
function ensureCalc(): void {
  if (!session || calcSent) return;
  calcSent = true;
  post({ type: 'calc', report: session.ensureEngine() });
}

scope.onmessage = (e) => {
  const req = e.data;
  try {
    switch (req.type) {
      case 'open': {
        session = WorkbookSession.open(new Uint8Array(req.bytes), req.fileName, req.locale, req.options);
        calcSent = false;
        post({ id: req.id, type: 'opened', result: session.loadResult() });
        // Let the grid render first; the engine builds in the next task.
        const opened = session;
        setTimeout(() => {
          if (session === opened) ensureCalc();
        }, 0);
        break;
      }
      case 'edit': {
        if (!session) throw new Error('No workbook is open.');
        ensureCalc();
        post({ id: req.id, type: 'edited', result: session.apply(req.op) });
        break;
      }
      case 'save': {
        if (!session) throw new Error('No workbook is open.');
        post({ id: req.id, type: 'saved', bytes: session.save(), fileName: session.saveFileName, changes: session.changes() });
        break;
      }
      case 'validate': {
        if (!session) throw new Error('No workbook is open.');
        ensureCalc();
        post({ id: req.id, type: 'validated', ok: session.checkValidation(req.sheet, req.row, req.col, req.text, req.rule) });
        break;
      }
      case 'listSource': {
        if (!session) throw new Error('No workbook is open.');
        ensureCalc();
        post({ id: req.id, type: 'listSource', source: session.listSource(req.sheet, req.formula) ?? null });
        break;
      }
      case 'changes': {
        if (!session) throw new Error('No workbook is open.');
        post({ id: req.id, type: 'changes', changes: session.changes() });
        break;
      }
    }
  } catch (err) {
    post({ id: req.id, type: 'error', message: err instanceof Error ? err.message : String(err) });
  }
};

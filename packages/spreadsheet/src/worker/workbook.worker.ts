import { errorText, msg, TextError, type Text } from '../i18n/message';
import { chooseEngine, chosenEngine } from '../core/engine';
import type { EngineId, EngineInfo } from '../core/engineIds';
import { WorkbookSession } from '../core/session';
import { heapBytes, loadEngine } from './engines';
import type { WorkerRequest, WorkerResponse } from './protocol';

/** The workbook, its parsed sheets and the calculation engine live here, off the UI thread. */
let session: WorkbookSession | undefined;
let calcSent = false;
/** Sessions before and after each structural change, for undo and redo; `at` is the current one. */
let sessions: WorkbookSession[] = [];
let at = 0;
/** How many structural changes can be undone; older states are let go to free memory. */
const STRUCTURE_UNDO_LIMIT = 20;
/** Each state kept for undo holds the whole workbook: for large ones, fewer are kept. */
const UNDO_CELL_BUDGET = 5_000_000;

const scope = self as unknown as {
  postMessage(message: WorkerResponse): void;
  onmessage: ((e: MessageEvent<WorkerRequest>) => void) | null;
};

function post(message: WorkerResponse): void {
  scope.postMessage(message);
}

/** Makes `next` the current session and sends the whole workbook to the UI. */
function adopt(next: WorkbookSession, id: number, warnings: Text[] = []): void {
  // Only the current state needs a calculation engine; one kept for undo builds it again if it is gone back to.
  if (session && session !== next) session.releaseEngine();
  session = next;
  const report = next.calcReport;
  calcSent = !!report;
  post({ id, type: 'restructured', result: next.loadResult(), updates: next.stateUpdates(), editedCount: next.editedCount, report, warnings });
  if (!report) calcSoon(next);
}

/** Builds the engine if needed and reports it, so its results always precede edit results. */
function ensureCalc(): void {
  if (!session || calcSent) return;
  calcSent = true;
  try {
    const started = performance.now();
    const report = session.ensureEngine();
    post({ type: 'calc', report, stats: { ms: Math.round(performance.now() - started), heapBytes: heapBytes() } });
  } catch (err) {
    // The session is left half calculated, so nothing more can be done with it here.
    post({ type: 'failed', text: errorText(err) });
    throw err;
  }
}

/** Builds the engine in the next task, so the grid can show first. */
function calcSoon(s: WorkbookSession): void {
  setTimeout(() => {
    if (session !== s) return;
    try {
      ensureCalc();
    } catch {
      // Reported as 'failed'.
    }
  }, 0);
}

/** While an engine loads for an open, later requests wait for it, in order; otherwise each is handled at once. */
let loading: Promise<void> | undefined;
scope.onmessage = (e) => {
  const req = e.data;
  const result = loading ? loading.then(() => handle(req)) : handle(req);
  if (!result) return;
  loading = result;
  void result.then(() => {
    if (loading === result) loading = undefined;
  });
};

function fail(id: number, err: unknown): void {
  post({ id, type: 'error', message: err instanceof Error ? err.message : String(err), text: errorText(err) });
}

function openWorkbook(req: Extract<WorkerRequest, { type: 'open' }>, engine: EngineInfo): void {
  session = WorkbookSession.open(new Uint8Array(req.bytes), req.fileName, req.locale, req.options);
  sessions = [session];
  at = 0;
  calcSent = false;
  post({ id: req.id, type: 'opened', result: session.loadResult(), engine });
  // Let the grid render first; the engine builds in the next task.
  calcSoon(session);
}

/** Handles a request; for an open that has to load its engine first, returns when it is done. */
function handle(req: WorkerRequest): Promise<void> | undefined {
  try {
    switch (req.type) {
      case 'open': {
        // Once the engine is loaded, opens go straight to it.
        const ready = chosenEngine();
        if (ready) {
          openWorkbook(req, { id: ready.id as EngineId, name: ready.name, loadMs: 0, transferBytes: 0, decodedBytes: 0 });
          break;
        }
        return loadEngine()
          .then(({ kind, info }) => {
            chooseEngine(kind);
            openWorkbook(req, info);
          })
          .catch((err: unknown) => fail(req.id, err));
      }
      case 'edit': {
        if (!session) throw new Error('No workbook is open.');
        ensureCalc();
        const started = performance.now();
        const result = session.apply(req.op);
        post({ id: req.id, type: 'edited', result, stats: { ms: Math.round((performance.now() - started) * 10) / 10, heapBytes: heapBytes() } });
        break;
      }
      case 'save': {
        if (!session) throw new Error('No workbook is open.');
        post({ id: req.id, type: 'saved', bytes: session.save(req.properties), fileName: session.saveFileName, changes: session.changes() });
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
      case 'conditionalFormats': {
        if (!session) throw new Error('No workbook is open.');
        ensureCalc();
        post({ id: req.id, type: 'conditionalFormats', result: session.conditionalFormats(req.sheet) });
        break;
      }
      case 'picture': {
        if (!session) throw new Error('No workbook is open.');
        // A copy: the package keeps its own for saving.
        const image = session.picture(req.part);
        post({ id: req.id, type: 'picture', image: image ? { bytes: image.bytes.slice(), type: image.type } : null });
        break;
      }
      case 'structure': {
        if (!session) throw new Error('No workbook is open.');
        const next = session.restructure(req.op);
        sessions = sessions.slice(0, at + 1);
        sessions.push(next);
        const limit = Math.max(1, Math.min(STRUCTURE_UNDO_LIMIT, Math.floor(UNDO_CELL_BUDGET / Math.max(1, next.cellCount))));
        while (sessions.length > limit + 1) sessions.shift();
        at = sessions.length - 1;
        adopt(next, req.id, next.structureWarnings);
        break;
      }
      case 'enableRecalc': {
        if (!session) throw new Error('No workbook is open.');
        calcSent = true;
        post({ id: req.id, type: 'recalcOn', report: session.enableRecalc() });
        break;
      }
      case 'undoStructure':
      case 'redoStructure': {
        const to = req.type === 'undoStructure' ? at - 1 : at + 1;
        if (!sessions[to]) throw new TextError(msg(req.type === 'undoStructure' ? 'core.structure.cantUndo' : 'core.structure.nothingToRedo'));
        at = to;
        adopt(sessions[at], req.id);
        break;
      }
      case 'changes': {
        if (!session) throw new Error('No workbook is open.');
        post({ id: req.id, type: 'changes', changes: session.changes() });
        break;
      }
    }
  } catch (err) {
    fail(req.id, err);
  }
  return undefined;
}

import { inRect } from '../core/refs';
import {
  CellFlag,
  type CalcReport,
  type DisplayStyle,
  type EditOp,
  type EditResult,
  type Finding,
  type Rect,
} from '../core/types';
import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import { displayText } from './format';
import { SheetView } from './sheetView';

export interface Toast {
  id: number;
  kind: 'info' | 'warn' | 'error';
  text: string;
}

export interface ConfirmRequest {
  title: string;
  body: string;
  confirmLabel: string;
  resolve: (ok: boolean) => void;
}

export interface Selection {
  /** Active cell, in sheet coordinates. */
  row: number;
  col: number;
  /** Selected ranges, in sheet coordinates. */
  rects: Rect[];
}

interface Clip {
  sheet: number;
  rect: Rect;
  first: string;
}

type Pending = { resolve: (r: WorkerResponse) => void; reject: (e: Error) => void };

const rectCells = (r: Rect) => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);

/**
 * Application state. The workbook itself lives in a Web Worker; this keeps
 * a view of each sheet for rendering, the undo history, and the UI state
 * around it (dialogs, toasts, panels).
 */
export class WorkbookStore {
  status: 'empty' | 'loading' | 'ready' | 'error' = 'empty';
  stage = '';
  error?: string;
  fileName = '';
  sheets: SheetView[] = [];
  styles: DisplayStyle[] = [];
  findings: Finding[] = [];
  calc?: CalcReport;
  date1904 = false;
  timings: Record<string, number> = {};
  active = 0;
  selections = new Map<number, Selection>();
  editedCount = 0;
  busy = false;
  undoStack: EditResult[] = [];
  redoStack: EditResult[] = [];
  toasts: Toast[] = [];
  confirmRequest?: ConfirmRequest;
  reportOpen = false;
  bannerDismissed = false;
  highlightEdits = true;
  version = 0;

  private worker?: Worker;
  private seq = 0;
  private pending = new Map<number, Pending>();
  private listeners = new Set<() => void>();
  private clip?: Clip;
  private queue: Promise<unknown> = Promise.resolve();
  private toastSeq = 0;

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getVersion = () => this.version;

  notify(): void {
    this.version++;
    for (const fn of this.listeners) fn();
  }

  get view(): SheetView | undefined {
    return this.sheets[this.active];
  }

  get selection(): Selection {
    return this.selections.get(this.active) ?? { row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] };
  }

  setSelection(sel: Selection): void {
    this.selections.set(this.active, sel);
    this.notify();
  }

  setActive(i: number): void {
    this.active = i;
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Worker plumbing

  private request(msg: Omit<WorkerRequest, 'id'> & { type: WorkerRequest['type'] }, transfer: Transferable[] = []): Promise<WorkerResponse> {
    const worker = this.worker;
    if (!worker) return Promise.reject(new Error('No workbook is open.'));
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      worker.postMessage({ ...msg, id } as WorkerRequest, transfer);
    });
  }

  private onMessage = (e: MessageEvent<WorkerResponse>) => {
    const msg = e.data;
    if (msg.type === 'calc') {
      this.calc = msg.report;
      this.findings = [...this.findings.filter((f) => f.kind !== 'formulaEngine'), ...msg.report.findings];
      for (const u of msg.report.updates) this.sheets[u.sheet]?.apply(u);
      this.notify();
      return;
    }
    const p = this.pending.get(msg.id);
    if (!p) return;
    this.pending.delete(msg.id);
    if (msg.type === 'error') p.reject(new Error(msg.message));
    else p.resolve(msg);
  };

  // ---------------------------------------------------------------------------
  // Opening and saving

  async open(fileName: string, bytes: ArrayBuffer): Promise<void> {
    if (this.editedCount > 0 && !(await this.confirm({
      title: 'Discard your changes?',
      body: `You have unsaved changes in ${this.fileName}. Opening another file discards them.`,
      confirmLabel: 'Discard and open',
    }))) {
      return;
    }
    if (!/\.xls[xm]$/i.test(fileName)) {
      this.toast('error', 'Only Excel workbooks (.xlsx or .xlsm) can be opened.');
      return;
    }
    this.worker?.terminate();
    for (const p of this.pending.values()) p.reject(new Error('Superseded'));
    this.pending.clear();
    // A fresh worker per file releases the previous workbook's memory.
    this.worker = new Worker(new URL('../worker/workbook.worker.ts', import.meta.url), { type: 'module' });
    this.worker.onmessage = this.onMessage;
    Object.assign(this, {
      status: 'loading',
      stage: `Reading ${fileName}…`,
      error: undefined,
      fileName,
      calc: undefined,
      undoStack: [],
      redoStack: [],
      editedCount: 0,
      bannerDismissed: false,
      reportOpen: false,
      clip: undefined,
    });
    this.selections.clear();
    this.notify();
    try {
      const res = await this.request(
        { type: 'open', bytes, fileName, locale: navigator.language || 'en-US' } as Omit<WorkerRequest, 'id'>,
        [bytes],
      );
      if (res.type !== 'opened') throw new Error('Unexpected response');
      const r = res.result;
      this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
      this.styles = r.styles;
      this.findings = r.findings;
      this.date1904 = r.date1904;
      this.timings = r.timings;
      this.active = Math.max(0, r.sheets.findIndex((s) => s.state === 'visible'));
      this.status = 'ready';
      this.stage = '';
    } catch (err) {
      if ((err as Error).message === 'Superseded') return;
      this.status = 'error';
      this.error = (err as Error).message;
    }
    this.notify();
  }

  async save(): Promise<{ bytes: Uint8Array; fileName: string } | undefined> {
    await this.queue;
    this.busy = true;
    this.notify();
    try {
      const res = await this.request({ type: 'save' } as Omit<WorkerRequest, 'id'>);
      if (res.type !== 'saved') return undefined;
      return { bytes: res.bytes, fileName: res.fileName };
    } catch (err) {
      this.toast('error', `Could not save: ${(err as Error).message}`);
      return undefined;
    } finally {
      this.busy = false;
      this.notify();
    }
  }

  async download(): Promise<void> {
    const saved = await this.save();
    if (!saved) return;
    const type = /\.xlsm$/i.test(saved.fileName)
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const url = URL.createObjectURL(new Blob([saved.bytes as BlobPart], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = saved.fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
    this.toast('info', this.editedCount ? `Saved ${saved.fileName} with ${this.editedCount.toLocaleString()} changed cells.` : `Saved ${saved.fileName} (unchanged).`);
  }

  // ---------------------------------------------------------------------------
  // Editing

  /** Runs an edit after any in-flight one, applies its results and records undo. */
  private edit(op: EditOp, record: 'undo' | 'none' = 'undo'): Promise<EditResult | undefined> {
    const run = async () => {
      this.busy = true;
      this.notify();
      try {
        const res = await this.request({ type: 'edit', op } as Omit<WorkerRequest, 'id'>);
        if (res.type !== 'edited') return undefined;
        const r = res.result;
        this.applyResult(r);
        if (record === 'undo' && r.undo.length) {
          this.undoStack.push(r);
          this.redoStack = [];
        }
        return r;
      } catch (err) {
        if ((err as Error).message !== 'Superseded') this.toast('error', (err as Error).message);
        return undefined;
      } finally {
        this.busy = false;
        this.notify();
      }
    };
    const p = this.queue.then(run, run);
    this.queue = p;
    return p;
  }

  private applyResult(r: EditResult): void {
    for (const s of r.newStyles) this.styles[s.index] = s.style;
    for (const u of r.updates) this.sheets[u.sheet]?.apply(u);
    this.editedCount = r.editedCount;
    if (r.rejected) this.toast('error', r.rejected);
    if (r.skipped) {
      const reason = r.skippedReasons[0] ?? '';
      this.toast('warn', r.skipped === 1 ? reason : `${r.skipped.toLocaleString()} cells were left unchanged. ${reason}`);
    }
    for (const w of r.warnings) this.toast('warn', w);
  }

  async undo(): Promise<void> {
    await this.queue;
    const entry = this.undoStack.pop();
    if (!entry) return;
    this.goTo(entry.undo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.undo }, 'none');
    this.redoStack.push(entry);
    this.notify();
  }

  async redo(): Promise<void> {
    await this.queue;
    const entry = this.redoStack.pop();
    if (!entry) return;
    this.goTo(entry.redo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.redo }, 'none');
    this.undoStack.push(entry);
    this.notify();
  }

  private goTo(sheet: number | undefined): void {
    if (sheet !== undefined && sheet !== this.active) this.setActive(sheet);
  }

  /** How many formulas in `rects` would be replaced by something that is not a formula. */
  private formulasLost(view: SheetView, rects: Rect[], keeps?: (row: number, col: number) => boolean): number {
    let n = 0;
    for (const rect of rects) {
      view.forEach(rect, (row, col, rec) => {
        if (!(rec.flags & CellFlag.Formula)) return;
        if (view.readOnlyReason(row, col)) return;
        if (keeps?.(row, col)) return;
        n++;
      });
    }
    return n;
  }

  private async okToReplace(n: number): Promise<boolean> {
    if (n === 0) return true;
    return this.confirm({
      title: n === 1 ? 'Replace a formula?' : `Replace ${n.toLocaleString()} formulas?`,
      body:
        n === 1
          ? 'This cell holds a formula. It will be replaced with a plain value, and the formula will not be in the saved file. You can undo this.'
          : 'Some of these cells hold formulas. They will be replaced with plain values, and those formulas will not be in the saved file. You can undo this.',
      confirmLabel: n === 1 ? 'Replace formula' : 'Replace formulas',
    });
  }

  async input(row: number, col: number, text: string): Promise<void> {
    const view = this.view;
    if (!view) return;
    const reason = view.readOnlyReason(row, col);
    if (reason) return this.toast('warn', reason);
    const rec = view.get(row, col);
    if (rec?.flags && rec.flags & CellFlag.Formula && !text.startsWith('=') && !(await this.okToReplace(1))) return;
    await this.edit({ type: 'input', sheet: this.active, row, col, text });
  }

  async clear(rects: Rect[]): Promise<void> {
    const view = this.view;
    if (!view) return;
    if (!(await this.okToReplace(this.formulasLost(view, rects)))) return;
    await this.edit({ type: 'clear', sheet: this.active, rects });
  }

  /** Records what was copied so a paste back into the workbook can carry formulas. */
  noteCopy(rect: Rect): void {
    const view = this.view;
    if (!view) return;
    const rec = view.get(rect.r1, rect.c1);
    const first = rec ? displayText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904).text : '';
    this.clip = { sheet: this.active, rect, first };
  }

  async paste(target: { row: number; col: number }, values: readonly (readonly string[])[], selection: Rect): Promise<void> {
    const view = this.view;
    if (!view || !values.length) return;
    const clip = this.clip;
    const h = values.length;
    const w = Math.max(...values.map((v) => v.length));
    const selH = selection.r2 - selection.r1 + 1;
    const selW = selection.c2 - selection.c1 + 1;
    // Like Excel, a selection that is a whole multiple of the copied block is filled with it.
    const tile = selH >= h && selW >= w && selH % h === 0 && selW % w === 0 && (selH > h || selW > w);
    const dst: Rect = tile ? selection : { r1: target.row, c1: target.col, r2: target.row + h - 1, c2: target.col + w - 1 };

    const internal =
      clip &&
      clip.rect.r2 - clip.rect.r1 + 1 === h &&
      clip.rect.c2 - clip.rect.c1 + 1 === w &&
      (values[0][0] ?? '') === clip.first;
    if (internal) {
      const src = this.sheets[clip.sheet];
      const keeps = (row: number, col: number) => {
        const sr = clip.rect.r1 + ((row - dst.r1) % h);
        const sc = clip.rect.c1 + ((col - dst.c1) % w);
        return !!((src.get(sr, sc)?.flags ?? 0) & CellFlag.Formula);
      };
      if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
      await this.edit({ type: 'copy', srcSheet: clip.sheet, src: clip.rect, sheet: this.active, dst });
      return;
    }

    let grid = values.map((line) => [...line]);
    if (tile) {
      grid = [];
      for (let r = 0; r < selH; r++) grid.push(Array.from({ length: selW }, (_, c) => values[r % h][c % w] ?? ''));
    }
    const keeps = (row: number, col: number) => (grid[row - dst.r1]?.[col - dst.c1] ?? '').startsWith('=');
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    if (rectCells(dst) > 200_000) this.toast('info', `Pasting ${rectCells(dst).toLocaleString()} cells…`);
    await this.edit({ type: 'pasteText', sheet: this.active, row: dst.r1, col: dst.c1, values: grid });
  }

  async fill(src: Rect, dst: Rect): Promise<void> {
    const view = this.view;
    if (!view) return;
    const h = src.r2 - src.r1 + 1;
    const w = src.c2 - src.c1 + 1;
    const mod = (a: number, n: number) => ((a % n) + n) % n;
    const keeps = (row: number, col: number) => {
      if (inRect(src, row, col)) return true;
      const sr = src.r1 + mod(row - src.r1, h);
      const sc = src.c1 + mod(col - src.c1, w);
      return !!((view.get(sr, sc)?.flags ?? 0) & CellFlag.Formula);
    };
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    await this.edit({ type: 'fill', sheet: this.active, src, dst });
  }

  // ---------------------------------------------------------------------------
  // Dialogs and messages

  confirm(req: Omit<ConfirmRequest, 'resolve'>): Promise<boolean> {
    return new Promise((resolve) => {
      this.confirmRequest = {
        ...req,
        resolve: (ok) => {
          this.confirmRequest = undefined;
          this.notify();
          resolve(ok);
        },
      };
      this.notify();
    });
  }

  toast(kind: Toast['kind'], text: string): void {
    const id = ++this.toastSeq;
    this.toasts = [...this.toasts.slice(-3), { id, kind, text }];
    this.notify();
    setTimeout(() => this.dismissToast(id), kind === 'error' ? 8000 : 5000);
  }

  dismissToast(id: number): void {
    this.toasts = this.toasts.filter((t) => t.id !== id);
    this.notify();
  }

  set<K extends 'reportOpen' | 'bannerDismissed' | 'highlightEdits'>(key: K, value: WorkbookStore[K]): void {
    this[key] = value;
    this.notify();
  }
}

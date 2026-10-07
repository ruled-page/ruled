import type { CalcReport, CellChange, EditOp, EditResult, LoadResult } from '../core/types';

export type WorkerRequest =
  | { id: number; type: 'open'; bytes: ArrayBuffer; fileName: string; locale: string }
  | { id: number; type: 'edit'; op: EditOp }
  | { id: number; type: 'save' }
  | { id: number; type: 'changes' };

export type WorkerResponse =
  | { id: number; type: 'opened'; result: LoadResult }
  /** Sent once, after the calculation engine is built; always before any edit result. */
  | { type: 'calc'; report: CalcReport }
  | { id: number; type: 'edited'; result: EditResult }
  | { id: number; type: 'saved'; bytes: Uint8Array; fileName: string; changes: CellChange[] }
  | { id: number; type: 'changes'; changes: CellChange[] }
  | { id: number; type: 'error'; message: string };

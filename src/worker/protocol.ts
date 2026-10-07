import type { OpenOptions } from '../core/session';
import type { CalcReport, CellChange, CfResult, EditOp, EditResult, ListSource, LoadResult } from '../core/types';

export type WorkerRequest =
  | { id: number; type: 'open'; bytes: ArrayBuffer; fileName: string; locale: string; options?: OpenOptions }
  | { id: number; type: 'edit'; op: EditOp }
  | { id: number; type: 'save' }
  | { id: number; type: 'changes' }
  | { id: number; type: 'validate'; sheet: number; row: number; col: number; text: string; rule: number }
  | { id: number; type: 'listSource'; sheet: number; formula: string }
  | { id: number; type: 'conditionalFormats'; sheet: number };

export type WorkerResponse =
  | { id: number; type: 'opened'; result: LoadResult }
  /** Sent once, after the calculation engine is built; always before any edit result. */
  | { type: 'calc'; report: CalcReport }
  | { id: number; type: 'edited'; result: EditResult }
  | { id: number; type: 'saved'; bytes: Uint8Array; fileName: string; changes: CellChange[] }
  | { id: number; type: 'changes'; changes: CellChange[] }
  | { id: number; type: 'validated'; ok: boolean }
  | { id: number; type: 'listSource'; source: ListSource | null }
  | { id: number; type: 'conditionalFormats'; result: CfResult }
  | { id: number; type: 'error'; message: string };

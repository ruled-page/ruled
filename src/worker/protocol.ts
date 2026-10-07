import type { StructuralOp } from '../core/refTransform';
import type { OpenOptions } from '../core/session';
import type { CalcReport, CellChange, CellUpdate, CfResult, EditOp, EditResult, ListSource, LoadResult } from '../core/types';

export type WorkerRequest =
  | { id: number; type: 'open'; bytes: ArrayBuffer; fileName: string; locale: string; options?: OpenOptions }
  | { id: number; type: 'edit'; op: EditOp }
  | { id: number; type: 'save'; properties?: Record<string, string> }
  | { id: number; type: 'changes' }
  | { id: number; type: 'validate'; sheet: number; row: number; col: number; text: string; rule: number }
  | { id: number; type: 'listSource'; sheet: number; formula: string }
  | { id: number; type: 'conditionalFormats'; sheet: number }
  /** Insert or delete rows or columns, or move cells, rewriting references. */
  | { id: number; type: 'structure'; op: StructuralOp }
  /** Go back to, or forward to, the workbook as it was around a structural change. */
  | { id: number; type: 'undoStructure' | 'redoStructure' }
  /** Turn live recalculation on for a workbook opened with it off. */
  | { id: number; type: 'enableRecalc' };

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
  | { id: number; type: 'recalcOn'; report: CalcReport }
  /**
   * After a structural change, or undoing one, the whole workbook as it now
   * is: what the file holds, the cells that differ from it, and the engine's
   * report if it is already built (otherwise a 'calc' message follows).
   */
  | { id: number; type: 'restructured'; result: LoadResult; updates: CellUpdate[]; editedCount: number; report?: CalcReport; warnings: string[] }
  | { id: number; type: 'error'; message: string };

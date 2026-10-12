import { engineFormula } from '@sitepen/formula-engine/formula';
import type { EngineKind, EngineValue, RawContent, Unsupported } from '@sitepen/formula-engine';
import { msg, type Message, type MessageKey, type Text } from '../i18n/message';
import type { CellContent } from './types';

/**
 * The editor's side of the calculation engine: which engine new sessions
 * use, cell contents as the engine takes them, and what isn't calculated or
 * was worked out, in the editor's words. The engine's own contract is in
 * "@sitepen/formula-engine".
 */

const LABELS: Record<Exclude<Unsupported['kind'], 'function'>, MessageKey> = {
  tableRefs: 'core.unsupported.tableRefs',
  externalLinks: 'core.unsupported.externalLinks',
  syntax: 'core.unsupported.syntax',
  array: 'core.unsupported.array',
  sharedFormula: 'core.unsupported.sharedFormula',
  tooLarge: 'core.unsupported.tooLarge',
};

const EXPLANATIONS: Record<Exclude<Unsupported['kind'], 'function'>, MessageKey> = {
  tableRefs: 'core.cellReason.tableRefs',
  externalLinks: 'core.cellReason.externalLinks',
  syntax: 'core.cellReason.syntax',
  array: 'core.cellReason.array',
  sharedFormula: 'core.cellReason.sharedFormula',
  tooLarge: 'core.cellReason.tooLarge',
};

/** What isn't calculated, as the recalculation report lists it: "LET function". */
export function unsupportedLabel(u: Unsupported): Message {
  return u.kind === 'function' ? msg('core.unsupported.function', { name: u.name }) : msg(LABELS[u.kind]);
}

/** Why a formula isn't calculated, in a sentence about its cell. */
export function unsupportedExplanation(u: Unsupported): Message {
  return u.kind === 'function' ? msg('core.cellReason.function', { name: u.name }) : msg(EXPLANATIONS[u.kind]);
}

export function toRaw(content: CellContent): RawContent {
  switch (content.kind) {
    case 'empty':
      return null;
    case 'number':
      return content.value;
    case 'bool':
      return content.value;
    // A leading apostrophe stops the engine from reading "007" or "1/2" as numbers.
    case 'string':
      return "'" + content.value;
    case 'error':
      return content.value;
    case 'formula':
      return engineFormula(content.formula);
  }
}

let fallback: EngineKind | undefined;

/** Sets the engine new sessions calculate with. */
export function chooseEngine(kind: EngineKind): void {
  fallback = kind;
}

/** The engine chosen for new sessions, if one has been. */
export function chosenEngine(): EngineKind | undefined {
  return fallback;
}

/** The engine new sessions calculate with. */
export function currentEngine(): EngineKind {
  if (!fallback) throw new Error('No calculation engine has been chosen: call chooseEngine first.');
  return fallback;
}

/** A value as a message shows it: a number, "text" in quotes, an error code, or "(empty)". */
export function describeValue(v: EngineValue): Text {
  if (v === null) return msg('core.value.empty');
  if (typeof v === 'object') return v.error;
  if (typeof v === 'string') return JSON.stringify(v);
  return String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'FALSE' ? String(v).toUpperCase() : String(v);
}

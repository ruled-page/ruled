import { useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent, type ReactNode, type Ref } from 'react';
import {
  argumentAt,
  completionAt,
  formulaReferences,
  insertFunction,
  signature,
  toggleAbsolute,
  type Completion,
  type FormulaRef,
} from '../core/formulaAssist';
import type { FunctionHelp } from '../core/functionCatalog';
import { useT } from '../i18n/react';

/**
 * Help while a formula is typed, shared by the formula bar and the cell
 * editor: the list of functions matching a name being typed, the arguments
 * of the function the caret is in, the references coloured, and F4.
 */
export interface FormulaAssist {
  /** The functions the name being typed could be, while the list is open. */
  completion?: Completion;
  /** The highlighted function in the list. */
  active: number;
  /** The function the caret is in, and which argument. */
  hint?: { fn: FunctionHelp; start: number; index: number };
  /** Where the hint or list is shown from: the start of the name being typed, or of the function. */
  anchor?: number;
  refs: FormulaRef[];
  /** Tells the help that the text changed by typing, which opens the list when a name is being typed. */
  typed(text: string, caret: number): void;
  dismiss(): void;
  setActive(index: number): void;
}

/** The help for a field holding `text` with the selection from `start` to `end`; off, it gives none. */
export function useFormulaAssist(text: string, start: number, end: number, on: boolean): FormulaAssist {
  const formula = on && text.startsWith('=');
  // The list is offered only straight after typing, as in Excel: moving the caret or Escape closes it.
  const [typedAt, setTypedAt] = useState<{ text: string; caret: number }>();
  const [picked, setPicked] = useState({ word: '', index: 0 });
  const caret = start === end ? end : -1;
  const completion = useMemo(
    () => (formula && typedAt?.text === text && typedAt.caret === caret ? completionAt(text, caret) : undefined),
    [formula, typedAt, text, caret],
  );
  const hint = useMemo(() => (formula ? argumentAt(text, end) : undefined), [formula, text, end]);
  const refs = useMemo(() => (formula ? formulaReferences(text) : []), [formula, text]);
  // Each new word starts at the top of its list.
  const word = completion ? text.slice(completion.start, completion.end) : '';
  const active = completion && picked.word === word ? Math.min(picked.index, completion.matches.length - 1) : 0;
  const anchor = completion?.start ?? hint?.start;
  return {
    completion,
    active,
    hint,
    anchor,
    refs,
    typed: (t, c) => setTypedAt({ text: t, caret: c }),
    dismiss: () => setTypedAt(undefined),
    setActive: (index) => setPicked({ word, index }),
  };
}

/** Inserts the function at `index` in the list in place of the name being typed, and closes the list. */
export function pickFunction(assist: FormulaAssist, text: string, index: number, apply: (text: string, start: number, end: number) => void): void {
  const c = assist.completion;
  if (!c) return;
  const done = insertFunction(text, c, c.matches[index].name);
  assist.dismiss();
  apply(done.text, done.caret, done.caret);
}

/**
 * The help's keys: in the list, the arrow keys move, Tab or Enter insert the
 * function and Escape closes it; F4 moves a reference between relative and
 * absolute. Returns whether the key was used; the caller then stops it.
 * `apply` puts new text and a selection into the field.
 */
export function assistKey(
  e: KeyboardEvent<HTMLElement>,
  assist: FormulaAssist,
  text: string,
  start: number,
  end: number,
  apply: (text: string, start: number, end: number) => void,
): boolean {
  if (e.nativeEvent.isComposing || e.altKey || e.ctrlKey || e.metaKey) return false;
  const c = assist.completion;
  if (c) {
    const n = c.matches.length;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      assist.setActive((assist.active + (e.key === 'ArrowDown' ? 1 : n - 1)) % n);
      return true;
    }
    if ((e.key === 'Enter' || e.key === 'Tab') && !e.shiftKey) {
      pickFunction(assist, text, assist.active, apply);
      return true;
    }
    if (e.key === 'Escape') {
      assist.dismiss();
      return true;
    }
  }
  if (e.key === 'F4' && !e.shiftKey && text.startsWith('=')) {
    const moved = toggleAbsolute(text, start, end);
    if (moved) apply(moved.text, moved.start, moved.end);
    return true;
  }
  return false;
}

/**
 * The formula's text with its references in their colours, drawn over the
 * field the formula is typed in (whose own text is then transparent). An
 * empty marker at `anchor` shows where the help's popup belongs.
 */
export function FormulaText({ text, refs, anchor, anchorRef }: { text: string; refs: FormulaRef[]; anchor?: number; anchorRef?: Ref<HTMLSpanElement> }) {
  const cuts = new Set([0, text.length, ...(anchor === undefined ? [] : [anchor]), ...refs.flatMap((r) => [r.start, r.end])]);
  const at = [...cuts].sort((a, b) => a - b);
  const out: ReactNode[] = [];
  for (let i = 0; i < at.length; i++) {
    if (at[i] === anchor) out.push(<span key="anchor" ref={anchorRef} />);
    if (i === at.length - 1) break;
    const piece = text.slice(at[i], at[i + 1]);
    const ref = refs.find((r) => r.start <= at[i] && at[i] < r.end);
    out.push(
      ref ? (
        <span key={at[i]} className={`wbe-ref wbe-ref-${ref.color + 1}`}>
          {piece}
        </span>
      ) : (
        piece
      ),
    );
  }
  return <>{out}</>;
}

/** The list of functions matching the name being typed, and what the highlighted one does. */
export function FunctionList({ id, completion, active, onPick, onHover }: { id: string; completion: Completion; active: number; onPick: (i: number) => void; onHover: (i: number) => void }) {
  const t = useT();
  const list = useRef<HTMLUListElement>(null);
  useEffect(() => {
    list.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
  }, [active, completion]);
  const fn = completion.matches[active];
  return (
    <>
      <ul
        className="wbe-fx-list"
        id={id}
        role="listbox"
        aria-label={t('formulaHelp.functions')}
        ref={list}
        // The list can't be scrolled by the user directly (see editor.css); the wheel scrolls it here.
        onWheel={(e) => {
          e.currentTarget.scrollTop += e.deltaY * (e.deltaMode === 1 ? 23 : 1);
        }}
      >
        {completion.matches.map((f, i) => (
          <li
            key={f.name}
            id={`${id}-${i}`}
            role="option"
            aria-selected={i === active}
            className={i === active ? 'wbe-active' : ''}
            onMouseEnter={() => onHover(i)}
            onMouseDown={(e) => {
              // Keeps the focus, and the caret, in the formula.
              e.preventDefault();
              onPick(i);
            }}
          >
            <span>{f.name}</span>
            {!f.calculated && <span className="wbe-fx-tag">{t('formulaHelp.notCalculated')}</span>}
          </li>
        ))}
      </ul>
      {fn?.args !== undefined && (
        <div className="wbe-fx-about">
          <div className="wbe-fx-signature">{signature(fn, -1).map((p) => p.text)}</div>
          {fn.description && <div>{fn.description}</div>}
        </div>
      )}
    </>
  );
}

/** The signature of the function the caret is in, the argument being typed in bold. */
export function ArgumentHint({ id, fn, index }: { id: string; fn: FunctionHelp; index: number }) {
  const t = useT();
  return (
    <div className="wbe-fx-hint" id={id} role="tooltip">
      <span className="wbe-fx-signature">
        {signature(fn, index).map((p, i) => (p.current ? <b key={i}>{p.text}</b> : p.text))}
      </span>
      {!fn.calculated && <span className="wbe-fx-tag">{t('formulaHelp.notCalculated')}</span>}
    </div>
  );
}

/** Says the function and argument the caret is in to screen readers, as it changes. */
export function HintAnnouncer({ hint }: { hint?: { fn: FunctionHelp; index: number } }) {
  const t = useT();
  const argument = hint && signature(hint.fn, hint.index).find((p) => p.current)?.text;
  const name = hint?.fn.name;
  return (
    <span className="wbe-sr-only" aria-live="polite">
      {name && (argument ? t('formulaHelp.spokenArgument', { name, number: hint.index + 1, argument }) : t('formulaHelp.spokenFunction', { name }))}
    </span>
  );
}

/** The popup under the field: the list while a name is typed, else the hint. */
export function AssistPopup({ assist, listId, hintId, style, className, onPick }: { assist: FormulaAssist; listId: string; hintId: string; style?: CSSProperties; className?: string; onPick: (i: number) => void }) {
  const { completion, hint } = assist;
  if (!completion && !hint) return null;
  return (
    <div className={`wbe-fx-popup ${className ?? ''}`} style={style}>
      {completion ? (
        <FunctionList id={listId} completion={completion} active={assist.active} onPick={onPick} onHover={assist.setActive} />
      ) : (
        hint && <ArgumentHint id={hintId} fn={hint.fn} index={hint.index} />
      )}
    </div>
  );
}

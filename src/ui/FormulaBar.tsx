import { useEffect, useRef, useState } from 'react';
import { rangeToA1, toA1 } from '../core/refs';
import { CellFlag, type CellNote } from '../core/types';
import { describeRule } from '../core/dataValidation';
import { displayText, editText } from '../state/format';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Icon } from './icons';
import type { Translator } from '../i18n';
import { msg } from '../i18n/message';
import { useT } from '../i18n/react';

type Tone = 'info' | 'warn' | 'lock' | 'edit' | 'error' | 'field' | 'note';

/** One line for a cell's note or comment thread. */
export function noteSummary(t: Translator, note: CellNote): string {
  const [first, ...replies] = note.entries;
  const text = (first?.text ?? '').replace(/\s+/g, ' ').trim();
  const author = first?.author;
  const thread = note.kind === 'thread';
  const summary = author
    ? t(thread ? 'formulaBar.commentBy' : 'formulaBar.noteBy', { author, text })
    : t(thread ? 'formulaBar.comment' : 'formulaBar.note', { text });
  return replies.length ? t('formulaBar.withReplies', { summary, count: replies.length }) : summary;
}

/** Excel's input message for a cell: its title, if it has one, and its text. */
function prompt(t: Translator, title: string | undefined, text: string): string {
  return title ? t('formulaBar.titledPrompt', { title, prompt: text }) : text;
}

function cellNote(t: Translator, store: WorkbookStore, view: SheetView, sel: Selection): { tone: Tone; text: string } | null {
  const readOnly = view.readOnlyReason(sel.row, sel.col);
  if (readOnly) return { tone: 'lock', text: t.text(readOnly) };
  const ctl = store.schema;
  const issues = ctl?.cellIssues(store.active, sel.row, sel.col) ?? [];
  if (issues.length) {
    return { tone: issues.some((i) => i.severity === 'error') ? 'error' : 'warn', text: issues.map((i) => t.text(i.message)).join(' ') };
  }
  if (ctl?.isPending(store.active, sel.row, sel.col)) return { tone: 'info', text: t('formulaBar.checking') };
  const described = ctl?.describe(store.active, sel.row, sel.col);
  let describe = described ? t.text(described) : undefined;
  const dv = describe ? undefined : store.listAt(store.active, sel.row, sel.col);
  if (dv) {
    const options = store.listOptions(dv, sel.row, sel.col);
    describe = dv.prompt
      ? prompt(t, dv.promptTitle, dv.prompt)
      : options
        ? t('formulaBar.chooseFromList', { options: `${options.slice(0, 8).join(', ')}${options.length > 8 ? '…' : ''}` })
        : t('formulaBar.restrictedList');
  } else if (!describe) {
    const rule = store.ruleAt(store.active, sel.row, sel.col)?.rule;
    if (rule) describe = rule.prompt ? prompt(t, rule.promptTitle, rule.prompt) : t.text(describeRule(rule));
  }
  const rec = view.get(sel.row, sel.col);
  const comment = view.noteAt(sel.row, sel.col);
  if (!rec) return describe ? { tone: 'field', text: describe } : comment ? { tone: 'note', text: noteSummary(t, comment) } : null;
  const reason = store.calc?.cellReasons.get(`${store.active}:${sel.row}:${sel.col}`);
  const why = reason ? t.text(reason) : undefined;
  if (rec.flags & CellFlag.Stale) {
    return { tone: 'warn', text: why ? t('formulaBar.staleBecause', { reason: why }) : t('formulaBar.stale') };
  }
  if (rec.flags & CellFlag.Edited && rec.flags & CellFlag.Unsupported) {
    return { tone: 'info', text: t(store.calc?.off ? 'formulaBar.recalcOff' : 'formulaBar.cantCalculateEdited') };
  }
  if (rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch)) {
    return { tone: 'info', text: t('formulaBar.savedValue', { reason: why ?? t('formulaBar.cantCalculate') }) };
  }
  if (describe) return { tone: 'field', text: describe };
  if (comment) return { tone: 'note', text: noteSummary(t, comment) };
  if (rec.flags & CellFlag.Edited) return { tone: 'edit', text: t('formulaBar.edited') };
  return null;
}

/** Shows the selection; type a cell, range or Sheet!ref and press Enter to go there. */
function NameBox({ store, name }: { store: WorkbookStore; name: string }) {
  const t = useT();
  const [draft, setDraft] = useState<string | null>(null);
  // Set while a click is focusing it, so the click's mouseup doesn't undo the select-all.
  const clicked = useRef(false);
  return (
    <input
      className="name-box"
      value={draft ?? name}
      spellCheck={false}
      aria-label={t('formulaBar.goTo')}
      title={t('formulaBar.goToTitle')}
      onMouseDown={(e) => {
        clicked.current = document.activeElement !== e.currentTarget;
      }}
      onMouseUp={(e) => {
        if (clicked.current) e.preventDefault();
        clicked.current = false;
      }}
      onFocus={(e) => {
        setDraft(name);
        e.target.select();
      }}
      onBlur={() => setDraft(null)}
      onChange={(e) => setDraft(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          if (store.goTo(draft ?? name)) setDraft(null);
          else store.toast('warn', msg('formulaBar.notARef', { ref: draft }));
        }
        if (e.key === 'Escape') {
          setDraft(null);
          e.currentTarget.blur();
        }
      }}
    />
  );
}

export function FormulaBar({ store, view, selection }: { store: WorkbookStore; view: SheetView; selection: Selection }) {
  const t = useT();
  const rec = view.get(selection.row, selection.col);
  const style = store.styles[rec?.style ?? 0] ?? store.styles[0];
  // A protected sheet can hide a cell's formula from the formula bar, as Excel does; its value still shows.
  const hidden = !!rec && rec.f !== undefined && view.info.protected && !!style.formulaHidden;
  const value = hidden ? displayText(rec, style, store.date1904).text : editText(rec, style, store.date1904, store.locale);
  const [draft, setDraft] = useState(value);
  const [focused, setFocused] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const readOnly = !!view.readOnlyReason(selection.row, selection.col);

  useEffect(() => {
    if (!focused) setDraft(value);
  }, [value, focused, selection.row, selection.col]);

  const rect = selection.rects[0];
  const name = rect && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) ? rangeToA1(rect) : toA1(selection.row, selection.col);
  const note = cellNote(t, store, view, selection);

  const commit = () => {
    if (draft !== value) void store.input(selection.row, selection.col, draft);
    input.current?.blur();
  };

  return (
    <div className="formula-bar">
      <NameBox store={store} name={name} />
      <div className={`fx ${rec?.f !== undefined ? 'is-formula' : ''}`}>
        <Icon name="formula" size={15} />
      </div>
      <input
        ref={input}
        className="formula-input"
        value={draft}
        readOnly={readOnly}
        spellCheck={false}
        aria-label={t('formulaBar.contents')}
        onFocus={() => setFocused(true)}
        onBlur={() => setFocused(false)}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') {
            setDraft(value);
            input.current?.blur();
          }
        }}
      />
      {note && (
        <div className={`cell-note tone-${note.tone}`} title={note.text}>
          <Icon name={note.tone === 'lock' ? 'lock' : note.tone === 'edit' ? 'edit' : note.tone === 'warn' || note.tone === 'error' ? 'warn' : note.tone === 'field' ? 'schema' : note.tone === 'note' ? 'note' : 'info'} size={14} />
          <span>{note.text}</span>
        </div>
      )}
      {/* The right-click menu, for touch screens (and anyone who looks for it here). */}
      <button
        className="cell-menu-btn"
        aria-label={t('formulaBar.cellMenu')}
        title={t('formulaBar.cellMenu')}
        aria-haspopup="menu"
        onClick={(e) => {
          const r = e.currentTarget.getBoundingClientRect();
          store.set('menu', { x: r.right, y: r.bottom, kind: 'cell' });
        }}
      >
        <Icon name="more" />
      </button>
    </div>
  );
}

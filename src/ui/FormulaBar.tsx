import { useEffect, useRef, useState } from 'react';
import { rangeToA1, toA1 } from '../core/refs';
import { CellFlag, type CellNote } from '../core/types';
import { describeRule } from '../core/dataValidation';
import { editText } from '../state/format';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Icon } from './icons';

type Tone = 'info' | 'warn' | 'lock' | 'edit' | 'error' | 'field' | 'note';

/** One line for a cell's note or comment thread. */
export function noteSummary(note: CellNote): string {
  const [first, ...replies] = note.entries;
  const who = first?.author ? `${first.author}: ` : '';
  const more = replies.length ? ` (${replies.length} ${replies.length === 1 ? 'reply' : 'replies'})` : '';
  return `${note.kind === 'thread' ? 'Comment' : 'Note'} — ${who}${(first?.text ?? '').replace(/\s+/g, ' ').trim()}${more}`;
}

function cellNote(store: WorkbookStore, view: SheetView, sel: Selection): { tone: Tone; text: string } | null {
  const readOnly = view.readOnlyReason(sel.row, sel.col);
  if (readOnly) return { tone: 'lock', text: readOnly };
  const ctl = store.schema;
  const issues = ctl?.cellIssues(store.active, sel.row, sel.col) ?? [];
  if (issues.length) {
    return { tone: issues.some((i) => i.severity === 'error') ? 'error' : 'warn', text: issues.map((i) => i.message).join(' ') };
  }
  if (ctl?.isPending(store.active, sel.row, sel.col)) return { tone: 'info', text: 'Checking…' };
  let describe = ctl?.describe(store.active, sel.row, sel.col);
  const dv = describe ? undefined : store.listAt(store.active, sel.row, sel.col);
  if (dv) {
    const options = store.listOptions(dv, sel.row, sel.col);
    describe = dv.prompt
      ? `${dv.promptTitle ? dv.promptTitle + ': ' : ''}${dv.prompt}`
      : options
        ? `Choose from the list: ${options.slice(0, 8).join(', ')}${options.length > 8 ? '…' : ''}`
        : 'Excel restricts this cell to a list.';
  } else if (!describe) {
    const rule = store.ruleAt(store.active, sel.row, sel.col)?.rule;
    if (rule) describe = rule.prompt ? `${rule.promptTitle ? rule.promptTitle + ': ' : ''}${rule.prompt}` : describeRule(rule);
  }
  const rec = view.get(sel.row, sel.col);
  const comment = view.noteAt(sel.row, sel.col);
  if (!rec) return describe ? { tone: 'field', text: describe } : comment ? { tone: 'note', text: noteSummary(comment) } : null;
  const why = store.calc?.cellReasons.get(`${store.active}:${sel.row}:${sel.col}`);
  if (rec.flags & CellFlag.Stale) {
    return {
      tone: 'warn',
      text: `${why ? why + ' ' : ''}Its inputs have changed, so the value shown is out of date. Excel recalculates it when the file is opened.`,
    };
  }
  if (rec.flags & CellFlag.Edited && rec.flags & CellFlag.Unsupported) {
    return store.calc?.off
      ? { tone: 'info', text: 'Live recalculation is off, so this formula isn\'t calculated here. It is saved as typed and Excel calculates it when the file is opened.' }
      : { tone: 'info', text: "This editor can't calculate this formula. It is saved as typed and Excel calculates it when the file is opened." };
  }
  if (rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch)) {
    return { tone: 'info', text: `${why ?? "This editor can't calculate this formula."} Showing Excel's saved value.` };
  }
  if (describe) return { tone: 'field', text: describe };
  if (comment) return { tone: 'note', text: noteSummary(comment) };
  if (rec.flags & CellFlag.Edited) return { tone: 'edit', text: 'Changed in this session.' };
  return null;
}

/** Shows the selection; type a cell, range or Sheet!ref and press Enter to go there. */
function NameBox({ store, name }: { store: WorkbookStore; name: string }) {
  const [draft, setDraft] = useState<string | null>(null);
  // Set while a click is focusing it, so the click's mouseup doesn't undo the select-all.
  const clicked = useRef(false);
  return (
    <input
      className="name-box"
      value={draft ?? name}
      spellCheck={false}
      aria-label="Go to cell"
      title="Selected cell. Type a reference and press Enter to go there."
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
          else store.toast('warn', `"${draft}" isn't a cell or range on a sheet in this workbook.`);
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
  const rec = view.get(selection.row, selection.col);
  const style = store.styles[rec?.style ?? 0] ?? store.styles[0];
  const value = editText(rec, style, store.date1904);
  const [draft, setDraft] = useState(value);
  const [focused, setFocused] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const readOnly = !!view.readOnlyReason(selection.row, selection.col);

  useEffect(() => {
    if (!focused) setDraft(value);
  }, [value, focused, selection.row, selection.col]);

  const rect = selection.rects[0];
  const name = rect && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) ? rangeToA1(rect) : toA1(selection.row, selection.col);
  const note = cellNote(store, view, selection);

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
        aria-label="Cell contents"
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
    </div>
  );
}

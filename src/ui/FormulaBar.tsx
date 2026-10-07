import { useEffect, useRef, useState } from 'react';
import { rangeToA1, toA1 } from '../core/refs';
import { CellFlag } from '../core/types';
import { editText } from '../state/format';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Icon } from './icons';

function cellNote(store: WorkbookStore, view: SheetView, sel: Selection): { tone: 'info' | 'warn' | 'lock' | 'edit'; text: string } | null {
  const readOnly = view.readOnlyReason(sel.row, sel.col);
  if (readOnly) return { tone: 'lock', text: readOnly };
  const rec = view.get(sel.row, sel.col);
  if (!rec) return null;
  const why = store.calc?.cellReasons.get(`${store.active}:${sel.row}:${sel.col}`);
  if (rec.flags & CellFlag.Stale) {
    return {
      tone: 'warn',
      text: `${why ? why + ' ' : ''}Its inputs have changed, so the value shown is out of date. Excel recalculates it when the file is opened.`,
    };
  }
  if (rec.flags & CellFlag.Edited && rec.flags & CellFlag.Unsupported) {
    return { tone: 'info', text: "This editor can't calculate this formula. It is saved as typed and Excel calculates it when the file is opened." };
  }
  if (rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch)) {
    return { tone: 'info', text: `${why ?? "This editor can't calculate this formula."} Showing Excel's saved value.` };
  }
  if (rec.flags & CellFlag.Edited) return { tone: 'edit', text: 'Changed in this session.' };
  return null;
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
      <div className="name-box" title="Selected cell">
        {name}
      </div>
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
          <Icon name={note.tone === 'lock' ? 'lock' : note.tone === 'edit' ? 'edit' : note.tone === 'warn' ? 'warn' : 'info'} size={14} />
          <span>{note.text}</span>
        </div>
      )}
    </div>
  );
}

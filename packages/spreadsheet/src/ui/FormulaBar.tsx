import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { rangeToA1, toA1 } from '@ruled/formula/refs';
import { CellFlag, VT, type CellNote } from '../core/types';
import { describeRule } from '../core/dataValidation';
import { displayFormula } from '@ruled/formula/formula';
import { linkAddress } from '../core/links';
import { displayText, editText } from '../state/format';
import { nameBoxEnter, nameBoxNames, selectionName } from '../state/names';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { AssistPopup, assistKey, FormulaText, HintAnnouncer, pickFunction, useAutoComplete, useFormulaAssist } from './FormulaAssist';
import { Icon } from './icons';
import type { Translator } from '../i18n';
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
    // Said as well as coloured, since the same icon marks both.
    const text = issues.map((i) => t(i.severity === 'error' ? 'issue.error' : 'issue.warning', { message: t.text(i.message) })).join(' ');
    return { tone: issues.some((i) => i.severity === 'error') ? 'error' : 'warn', text };
  }
  if (ctl?.isPending(store.active, sel.row, sel.col)) return { tone: 'info', text: t('formulaBar.checking') };
  const described = ctl?.describe(store.active, sel.row, sel.col);
  let describe = described ? t.text(described) : undefined;
  const dv = describe ? undefined : store.listAt(store.active, sel.row, sel.col);
  if (dv) {
    const options = store.listOptions(dv, sel.row, sel.col);
    describe = dv.prompt && dv.showPrompt !== false
      ? prompt(t, dv.promptTitle, dv.prompt)
      : options
        ? t('formulaBar.chooseFromList', { options: `${options.slice(0, 8).join(', ')}${options.length > 8 ? '…' : ''}` })
        : t('formulaBar.restrictedList');
  } else if (!describe) {
    const rule = store.ruleAt(store.active, sel.row, sel.col)?.rule;
    if (rule) describe = rule.prompt && rule.showPrompt !== false ? prompt(t, rule.promptTitle, rule.prompt) : t.text(describeRule(rule)) || undefined;
  }
  const rec = view.get(sel.row, sel.col);
  const comment = view.noteAt(sel.row, sel.col);
  if (!rec) return describe ? { tone: 'field', text: describe } : comment ? { tone: 'note', text: noteSummary(t, comment) } : null;
  const reason = store.calc?.cellReasons.get(`${store.active}:${sel.row}:${sel.col}`);
  const why = reason ? t.text(reason) : undefined;
  const spill = rec.spill;
  // A formula whose results have no room, and the cells it needs.
  if (spill && rec.f !== undefined && rec.t === VT.Error && rec.s === '#SPILL!') return { tone: 'warn', text: t('formulaBar.spillBlocked', { range: rangeToA1(spill) }) };
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
  if (spill && rec.flags & CellFlag.Spilled) return { tone: 'info', text: t('formulaBar.spilled', { ref: toA1(spill.r1, spill.c1) }) };
  if (rec.flags & CellFlag.Edited) return { tone: 'edit', text: t('formulaBar.edited') };
  return null;
}

/**
 * Shows the selection, or its name; type a cell, range, Sheet!ref or name
 * and press Enter to go there, or a new name to name the selection. As in
 * Excel, its list offers the names that refer to cells: typing narrows it,
 * the arrow keys move through it (Alt+Down opens it), Enter picks.
 */
function NameBox({ store, name }: { store: WorkbookStore; name: string }) {
  const t = useT();
  const id = useId();
  const box = useRef<HTMLInputElement>(null);
  const [draft, setDraft] = useState<string | null>(null);
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(-1);
  // Set while a click is focusing it, so the click's mouseup doesn't undo the select-all.
  const clicked = useRef(false);
  const typed = draft !== null && draft !== name ? draft.trim().toUpperCase() : '';
  const all = nameBoxNames(store);
  const options = typed ? all.filter((n) => n.name.toUpperCase().startsWith(typed)) : all;
  const shown = open && options.length > 0;
  const close = () => {
    setOpen(false);
    setActive(-1);
  };
  const go = async (text: string) => {
    close();
    if (await nameBoxEnter(store, text)) setDraft(null);
  };
  return (
    <div className="wbe-name-box-wrap">
      <input
        ref={box}
        className="wbe-name-box"
        value={draft ?? name}
        spellCheck={false}
        autoComplete="off"
        role="combobox"
        aria-label={t('formulaBar.goTo')}
        aria-expanded={shown}
        aria-controls={`${id}-names`}
        aria-autocomplete="list"
        aria-activedescendant={shown && active >= 0 ? `${id}-${active}` : undefined}
        title={t('names.box.title')}
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
        onBlur={() => {
          setDraft(null);
          close();
        }}
        onChange={(e) => {
          setDraft(e.target.value);
          setOpen(e.target.value.trim() !== '');
          setActive(-1);
        }}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            e.preventDefault();
            if (e.altKey) {
              if (e.key === 'ArrowDown') setOpen(true);
              else close();
            } else if (!shown) {
              setOpen(true);
              setActive(e.key === 'ArrowDown' ? 0 : options.length - 1);
            } else {
              const step = e.key === 'ArrowDown' ? 1 : -1;
              setActive((i) => (i < 0 && step < 0 ? options.length - 1 : (i + step + options.length) % options.length));
            }
          } else if (e.key === 'Enter') {
            e.preventDefault();
            void go(shown && active >= 0 ? options[active].name : (draft ?? name));
          } else if (e.key === 'Escape') {
            // The first Escape closes the list; the next leaves the box as it was.
            if (shown) {
              e.stopPropagation();
              close();
            } else {
              setDraft(null);
              store.focusGrid();
            }
          }
        }}
      />
      <button
        className="wbe-name-box-toggle"
        tabIndex={-1}
        aria-label={t('names.box.show')}
        title={t('names.box.show')}
        // Keeps the focus in the box, whose list this opens.
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => {
          box.current?.focus();
          setDraft(name);
          if (shown) close();
          else setOpen(true);
        }}
      >
        <Icon name="chevronDown" size={14} />
      </button>
      <ul id={`${id}-names`} className="wbe-name-box-list" role="listbox" aria-label={t('names.box.list')} hidden={!shown}>
        {shown &&
          options.map((o, i) => (
            <li
              key={o.name}
              id={`${id}-${i}`}
              role="option"
              aria-selected={i === active}
              className={i === active ? 'wbe-active' : ''}
              onMouseEnter={() => setActive(i)}
              onMouseDown={(e) => {
                e.preventDefault();
                void go(o.name);
              }}
            >
              <span className="wbe-name-box-name">{o.name}</span>{' '}
              {/* What it refers to, as the Name Manager shows it, so a name on another sheet says where it goes. */}
              <span className="wbe-name-box-ref">{o.refersTo}</span>
            </li>
          ))}
      </ul>
    </div>
  );
}

export function FormulaBar({ store, view, selection }: { store: WorkbookStore; view: SheetView; selection: Selection }) {
  const t = useT();
  const rec = view.get(selection.row, selection.col);
  const style = store.styles[rec?.style ?? 0] ?? store.styles[0];
  // Each cell of an array formula entered with Ctrl+Shift+Enter shows the formula in braces, as in Excel.
  const array = rec?.array ? view.get(rec.array.r1, rec.array.c1) : null;
  const formulaStyle = array ? (store.styles[array.style] ?? store.styles[0]) : style;
  // A protected sheet can hide a cell's formula from the formula bar, as Excel does; its value still shows.
  const hidden = !!rec && (rec.f !== undefined || !!array) && view.info.protected && !!formulaStyle.formulaHidden;
  const value = hidden
    ? displayText(rec, style, store.date1904, false, store.locale).text
    : array?.f !== undefined
      ? `{${displayFormula(array.f)}}`
      : editText(rec, style, store.date1904, store.locale);
  // A cell a formula's results spill into holds nothing itself; the formula shows greyed, as in Excel.
  const spiller = rec?.spill && rec.flags & CellFlag.Spilled ? view.get(rec.spill.r1, rec.spill.c1) : null;
  const ghostStyle = store.styles[spiller?.style ?? 0] ?? store.styles[0];
  const ghost = spiller?.f !== undefined && !(view.info.protected && ghostStyle.formulaHidden) ? displayFormula(spiller.f) : undefined;
  const [draft, setDraft] = useState(value);
  const [focused, setFocused] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const readOnly = !!view.readOnlyReason(selection.row, selection.col);

  useEffect(() => {
    if (!focused) setDraft(value);
  }, [value, focused, selection.row, selection.col]);

  // Help while a formula is typed: its references coloured, functions offered, argument hints and F4.
  const id = useId();
  const field = useRef<HTMLDivElement>(null);
  const anchor = useRef<HTMLSpanElement>(null);
  const [sel, setSel] = useState<[number, number]>([0, 0]);
  const [scroll, setScroll] = useState(0);
  const [popupX, setPopupX] = useState(0);
  const pendingSel = useRef<[number, number]>(undefined);
  const assist = useFormulaAssist(store, draft, sel[0], sel[1], focused && !readOnly);
  const auto = useAutoComplete(store, selection.row, selection.col, focused && !readOnly && store.features.autoComplete);
  const colored = focused && draft.startsWith('=');
  const apply = (text: string, start: number, end: number) => {
    setDraft(text);
    setSel([start, end]);
    pendingSel.current = [start, end];
  };

  // The grid outlines the references of a formula typed here.
  useEffect(() => {
    if (focused) store.showFormulaRefs(draft);
  }, [store, focused, draft]);
  useEffect(() => {
    if (focused) return () => store.showFormulaRefs(undefined);
  }, [store, focused]);

  // Once drawn: the caret where the help put it, the coloured text scrolled
  // with the field's own, and the popup under the name it is about.
  useLayoutEffect(() => {
    const el = input.current;
    if (!el) return;
    if (pendingSel.current) {
      el.setSelectionRange(...pendingSel.current);
      pendingSel.current = undefined;
    }
    if (el.scrollLeft !== scroll) setScroll(el.scrollLeft);
    const at = anchor.current?.getBoundingClientRect().left;
    const box = field.current?.getBoundingClientRect();
    if (at !== undefined && box) {
      const x = Math.round(Math.max(0, Math.min(at - box.left, box.width - 280)));
      setPopupX((prev) => (prev === x ? prev : x));
    }
  }, [draft, sel, scroll, assist.anchor]);

  // Whole rows or columns are named by the active cell, as in Excel: their range stops at the used part.
  const rect = selection.lines ? undefined : selection.rects[0];
  const name = (rect && selectionName(store)) ?? (rect && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) ? rangeToA1(rect) : toA1(selection.row, selection.col));
  const note = cellNote(t, store, view, selection);
  const link = store.linkAt(selection.row, selection.col);
  const address = link && linkAddress(link);

  // Back to the grid, at the cell, as the keyboard was before the formula bar took it.
  const commit = () => {
    const text = auto.final(draft);
    if (text !== value) void store.input(selection.row, selection.col, text);
    store.focusGrid();
  };

  return (
    <div className="wbe-formula-bar">
      <NameBox store={store} name={name} />
      <div className={`wbe-fx ${rec?.f !== undefined ? 'wbe-is-formula' : ''}`}>
        <Icon name="formula" size={15} />
      </div>
      <div className={`wbe-formula-field ${colored ? 'wbe-colored' : ''}`} ref={field}>
        <input
          ref={input}
          className="wbe-formula-input"
          value={draft}
          placeholder={ghost}
          readOnly={readOnly}
          spellCheck={false}
          role="combobox"
          aria-label={t('formulaBar.contents')}
          aria-autocomplete="list"
          aria-expanded={!!assist.completion}
          aria-controls={assist.completion ? `${id}-list` : undefined}
          aria-activedescendant={assist.completion ? `${id}-list-${assist.active}` : undefined}
          aria-describedby={[assist.hint && !assist.completion && `${id}-hint`, note && `${id}-note`].filter(Boolean).join(' ') || undefined}
          onFocus={() => setFocused(true)}
          onBlur={() => {
            setFocused(false);
            assist.dismiss();
          }}
          onChange={(e) => {
            const start = e.target.selectionStart ?? e.target.value.length;
            const completed = auto.changed(e.target.value, start, e.target.selectionEnd ?? start, (e.nativeEvent as InputEvent).inputType);
            if (completed) {
              apply(completed.text, ...completed.sel);
            } else {
              setDraft(e.target.value);
              setSel([start, e.target.selectionEnd ?? start]);
            }
            assist.typed(e.target.value, start);
          }}
          onSelect={(e) => setSel([e.currentTarget.selectionStart ?? 0, e.currentTarget.selectionEnd ?? 0])}
          onScroll={(e) => setScroll(e.currentTarget.scrollLeft)}
          onKeyDown={(e) => {
            const el = e.currentTarget;
            if (assistKey(e, assist, draft, el.selectionStart ?? 0, el.selectionEnd ?? 0, apply)) {
              e.preventDefault();
              return;
            }
            if (e.key === 'Enter') commit();
            if (e.key === 'Escape') {
              setDraft(value);
              store.focusGrid();
            }
          }}
        />
        {colored && (
          <div className="wbe-formula-text" aria-hidden="true">
            <span style={{ transform: `translateX(${-scroll}px)` }}>
              <FormulaText text={draft} refs={assist.refs} anchor={assist.anchor} anchorRef={anchor} />
            </span>
          </div>
        )}
        <AssistPopup assist={assist} listId={`${id}-list`} hintId={`${id}-hint`} style={{ left: popupX }} onPick={(i) => pickFunction(assist, draft, i, apply)} />
        <HintAnnouncer hint={assist.hint} />
      </div>
      {note && (
        <div id={`${id}-note`} className={`wbe-cell-note wbe-tone-${note.tone}`} title={note.text}>
          <Icon name={note.tone === 'lock' ? 'lock' : note.tone === 'edit' ? 'edit' : note.tone === 'warn' || note.tone === 'error' ? 'warn' : note.tone === 'field' ? 'schema' : note.tone === 'note' ? 'note' : 'info'} size={14} />
          <span>{note.text}</span>
        </div>
      )}
      {/* Where the cell's link goes, in full; pressing it follows the link, on touch screens too. */}
      {address &&
        (store.followsLinks ? (
          <button
            className="wbe-cell-link"
            title={address}
            aria-label={t('links.formulaBar.open', { address })}
            onClick={() => void store.followLink(selection.row, selection.col)}
          >
            <Icon name="link" size={14} />
            <span>{address}</span>
          </button>
        ) : (
          <div className="wbe-cell-link" title={address}>
            <Icon name="link" size={14} />
            <span>{address}</span>
          </div>
        ))}
      {/* The right-click menu, for touch screens (and anyone who looks for it here). */}
      <button
        className="wbe-cell-menu-btn"
        aria-label={t('formulaBar.cellMenu')}
        title={t('formulaBar.cellMenu')}
        aria-haspopup="menu"
        aria-expanded={store.menu?.kind === 'cell'}
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

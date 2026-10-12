import { useEffect, useMemo, useRef, useState } from 'react';
import { toA1 } from '@sitepen/formula-engine/refs';
import type { FindQuery, WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { shortcutKey } from './keys';
import { useDialog } from './useDialog';
import { msg } from '../i18n/message';
import { useT } from '../i18n/react';
import { Select } from './Select';

/** Find, and find and replace, over the sheet or the whole workbook. */
export function FindPanel({ store }: { store: WorkbookStore }) {
  const t = useT();
  const [text, setText] = useState('');
  const [replacement, setReplacement] = useState('');
  const [replacing, setReplacing] = useState(false);
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [lookIn, setLookIn] = useState<FindQuery['lookIn']>('values');
  const [scope, setScope] = useState<FindQuery['scope']>('sheet');
  const [at, setAt] = useState(-1);
  const input = useRef<HTMLInputElement>(null);
  const panel = useRef<HTMLDivElement>(null);
  const query: FindQuery = { text, matchCase, wholeCell, lookIn: replacing ? 'formulas' : lookIn, scope };
  // Matches are worked out again whenever the workbook or the search changes.
  const version = store.sheets.reduce((n, v) => n + v.version, 0) + store.active * 1e9;
  const hits = useMemo(() => store.find(query), [store, text, matchCase, wholeCell, lookIn, scope, replacing, version]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    input.current?.focus();
    input.current?.select();
  }, []);
  useEffect(() => setAt(-1), [text, matchCase, wholeCell, lookIn, scope]);

  const go = (step: 1 | -1) => {
    if (!hits.length) return;
    const sel = store.selection;
    // Start from the active cell, so "next" means the next one after it.
    let i = at;
    if (i < 0) {
      i = hits.findIndex((h) => h.sheet > store.active || (h.sheet === store.active && (h.row > sel.row || (h.row === sel.row && h.col > sel.col))));
      i = step > 0 ? (i < 0 ? 0 : i) : (i <= 0 ? hits.length : i) - 1;
    } else {
      i = (i + step + hits.length) % hits.length;
    }
    setAt(i);
    store.reveal(hits[i].sheet, hits[i].row, hits[i].col);
  };
  const replaceOne = async () => {
    const h = hits[at];
    if (!h) return go(1);
    await store.replace(query, replacement, h);
    go(1);
  };
  const replaceAll = async () => {
    const n = await store.replace(query, replacement);
    store.toast('info', n ? msg('find.replaced', { count: n }) : msg('find.nothingReplaced'));
  };
  const close = () => store.set('findOpen', false);
  // Escape closes it and the focus goes back to the grid; Tab can still leave it for the sheet.
  useDialog(panel, true, close, { initial: input });

  // ⌘G and ⇧⌘G (Ctrl+G elsewhere) find the next and previous match, as they do in the browser's own find.
  const step = useRef(go);
  step.current = go;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey) || e.altKey || shortcutKey(e) !== 'g') return;
      e.preventDefault();
      step.current(e.shiftKey ? -1 : 1);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);
  const current = hits[at];
  const where = current && (scope === 'workbook' ? `${store.sheets[current.sheet]?.info.name}!${toA1(current.row, current.col)}` : toA1(current.row, current.col));
  // Buttons that would do nothing stay focusable, so pressing one that runs out of matches keeps the focus.
  const none = !hits.length;

  return (
    <div
      ref={panel}
      className="wbe-find-panel"
      role="search"
      onKeyDown={(e) => {
        if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') {
          e.preventDefault();
          go(e.shiftKey ? -1 : 1);
        }
      }}
    >
      <div className="wbe-find-row">
        <button className="wbe-btn wbe-icon wbe-small" onClick={() => setReplacing(!replacing)} aria-label={t(replacing ? 'find.hideReplace' : 'find.showReplace')} title={t('find.replace')}>
          <Icon name={replacing ? 'minus' : 'plus'} size={14} />
        </button>
        <input ref={input} placeholder={t('find.find')} value={text} onChange={(e) => setText(e.target.value)} aria-label={t('find.find')} />
        <span className="wbe-find-count" aria-live="polite">
          {text ? (hits.length ? t('find.position', { at: at >= 0 ? t.number(at + 1) : '–', total: hits.length }) : t('find.noMatches')) : ''}
          {where && <span className="wbe-sr-only">{`, ${where}`}</span>}
        </span>
        <button className="wbe-btn wbe-small" onClick={() => go(-1)} aria-disabled={none} aria-label={t('find.previous')} title={t('find.previous')}>
          <span aria-hidden="true">↑</span>
        </button>
        <button className="wbe-btn wbe-small" onClick={() => go(1)} aria-disabled={none} aria-label={t('find.next')} title={t('find.next')}>
          <span aria-hidden="true">↓</span>
        </button>
        <button className="wbe-btn wbe-icon wbe-small" onClick={close} aria-label={t('find.close')}>
          <Icon name="close" size={14} />
        </button>
      </div>
      {replacing && (
        <div className="wbe-find-row">
          <span className="wbe-find-indent" />
          <input placeholder={t('find.replaceWith')} value={replacement} onChange={(e) => setReplacement(e.target.value)} aria-label={t('find.replaceWith')} />
          <button className="wbe-btn wbe-small" onClick={() => !none && void replaceOne()} aria-disabled={none}>
            {t('find.replace')}
          </button>
          <button className="wbe-btn wbe-small" onClick={() => !none && void replaceAll()} aria-disabled={none}>
            {t('find.replaceAll')}
          </button>
        </div>
      )}
      <div className="wbe-find-options">
        <label>
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> {t('find.matchCase')}
        </label>
        <label>
          <input type="checkbox" checked={wholeCell} onChange={(e) => setWholeCell(e.target.checked)} /> {t('find.wholeCell')}
        </label>
        {!replacing && (
          <Select value={lookIn} onChange={(e) => setLookIn(e.target.value as FindQuery['lookIn'])} aria-label={t('find.lookIn')}>
            <option value="values">{t('find.inValues')}</option>
            <option value="formulas">{t('find.inFormulas')}</option>
          </Select>
        )}
        <Select value={scope} onChange={(e) => setScope(e.target.value as FindQuery['scope'])} aria-label={t('find.within')}>
          <option value="sheet">{t('find.thisSheet')}</option>
          <option value="workbook">{t('find.workbook')}</option>
        </Select>
      </div>
      {current && scope === 'workbook' && <div className="wbe-find-where">{where}</div>}
    </div>
  );
}

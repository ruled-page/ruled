import { useId, useRef, useState } from 'react';
import { linkAddress, linkDestination, resolvePlace, splitPlace, typedAddress, type LinkFields } from '../core/links';
import { MAX_COLS, MAX_ROWS, parseRange, quoteSheetName, rangeToA1 } from '../core/refs';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';

/** A defined name in the place list, told apart from a sheet with the same name. */
const NAME = 'name:';

interface Draft {
  kind: 'web' | 'place';
  address: string;
  /** A sheet's name, or a defined name after NAME. */
  place: string;
  cell: string;
  text: string;
  tooltip: string;
}

/** What the dialog starts with: the active cell's link, if it has one of the sheet's, and its text. */
function draftFor(store: WorkbookStore): Draft {
  const view = store.view!;
  const { row, col } = store.selection;
  const own = store.linkAt()?.link;
  const draft: Draft = { kind: 'web', address: '', place: view.info.name, cell: 'A1', text: store.shownText(view, row, col), tooltip: own?.tooltip ?? '' };
  const to = own && linkDestination(own);
  if (to?.kind === 'place') {
    draft.kind = 'place';
    const at = splitPlace(to.location);
    const name = at && !at.sheet && !parseRange(at.ref) ? store.names.find((n) => n.name.toLowerCase() === at.ref.toLowerCase()) : undefined;
    if (name) draft.place = NAME + name.name;
    else if (at) {
      draft.place = store.sheets.find((v) => v.info.name.toLowerCase() === (at.sheet ?? view.info.name).toLowerCase())?.info.name ?? view.info.name;
      draft.cell = at.ref.replace(/\$/g, '');
    }
  } else if (own) {
    draft.address = linkAddress(own);
  }
  return draft;
}

/**
 * Gives the selected cells a link, or changes or takes away theirs, as
 * Excel's Insert Hyperlink dialog does: to a web page or email address, or to
 * a place in the workbook (a sheet's cells or a defined name), with the text
 * the cell shows and a screen tip.
 */
export function LinkDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const open = store.linkDialog;
  const dialog = useRef<HTMLFormElement>(null);
  const first = useRef<HTMLInputElement & HTMLSelectElement>(null);
  const cellInput = useRef<HTMLInputElement>(null);
  const id = useId();
  const [draft, setDraft] = useState<Draft | null>(null);
  const [tried, setTried] = useState(false);
  const opened = useRef(false);
  if (opened.current !== open) {
    opened.current = open;
    if (open && store.view) {
      setDraft(draftFor(store));
      setTried(false);
    }
  }
  const close = () => store.set('linkDialog', false);
  useDialog(dialog, open, close, { modal: true, initial: first });
  const view = store.view;
  if (!open || !draft || !view) return null;

  const set = (patch: Partial<Draft>) => setDraft({ ...draft, ...patch });
  const { row, col } = store.selection;
  const rect = store.selection.rects[0];
  const merge = view.mergeAt(row, col);
  const oneCell = !rect || (rect.r1 === rect.r2 && rect.c1 === rect.c2) || (!!merge && merge.r1 === rect.r1 && merge.c1 === rect.c1 && merge.r2 === rect.r2 && merge.c2 === rect.c2);
  // As in Excel, the text can't be changed over several cells, or in a cell that holds a formula.
  const textEditable = oneCell && view.get(row, col)?.f === undefined && !view.readOnlyReason(row, col);
  const editing = !!store.linkAt()?.link;
  const linked = store.linksInSelection();

  const sheets = store.sheets.filter((v) => v.info.kind === 'worksheet' && v.info.state === 'visible').map((v) => v.info.name);
  const names = store.names
    .filter((n) => !n.name.startsWith('_xlnm.') && (n.localSheet === undefined || n.localSheet === store.active))
    .filter((n) => resolvePlace(n.name, store.sheets.map((v) => v.info.name), store.names, store.active))
    .map((n) => n.name);
  const byName = draft.place.startsWith(NAME);
  const cells = parseRange(draft.cell.trim());

  const target = draft.kind === 'web' ? typedAddress(draft.address) : undefined;
  const addressError: MessageKey | undefined = draft.kind !== 'web' ? undefined : !draft.address.trim() ? 'links.dialog.addressMissing' : !target ? 'links.dialog.addressRefused' : undefined;
  const cellError: MessageKey | undefined = draft.kind === 'place' && !byName && (!cells || cells.r2 >= MAX_ROWS || cells.c2 >= MAX_COLS) ? 'links.dialog.cellInvalid' : undefined;
  const error = tried ? (addressError ?? cellError) : undefined;
  const describedBy = (hint: string, shown: boolean) => (shown ? `${id}-${hint} ${id}-error` : `${id}-${hint}`);

  const submit = () => {
    if (addressError || cellError) {
      setTried(true);
      (addressError ? first : cellInput).current?.focus();
      return;
    }
    const location = byName ? draft.place.slice(NAME.length) : `${quoteSheetName(draft.place)}!${rangeToA1(cells!)}`;
    const link: LinkFields = draft.kind === 'web' ? { target } : { location };
    // An empty cell shows the address, as Excel fills it in.
    const text = !textEditable ? undefined : draft.text.trim() ? draft.text : linkAddress(link);
    if (text) link.display = text;
    if (draft.tooltip.trim()) link.tooltip = draft.tooltip.trim();
    close();
    void store.setLink(link, text);
  };
  const remove = () => {
    close();
    void store.removeLinks();
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-link-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        noValidate
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          submit();
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t(editing ? 'links.dialog.editTitle' : 'links.dialog.insertTitle')}</h2>
          <label className="wbe-field">
            <span>{t('links.dialog.text')}</span>
            <input value={draft.text} disabled={!textEditable} onChange={(e) => set({ text: e.target.value })} aria-describedby={textEditable ? undefined : `${id}-text`} />
          </label>
          {!textEditable && (
            <p id={`${id}-text`} className="wbe-muted wbe-small">
              {t('links.dialog.textKept')}
            </p>
          )}
          <fieldset className="wbe-link-kind">
            <legend>{t('links.dialog.linkTo')}</legend>
            {(['web', 'place'] as const).map((kind) => (
              <label key={kind}>
                <input type="radio" name={`${id}-kind`} checked={draft.kind === kind} onChange={() => set({ kind })} />
                {t(kind === 'web' ? 'links.dialog.web' : 'links.dialog.place')}
              </label>
            ))}
          </fieldset>
          {draft.kind === 'web' ? (
            <>
              <label className="wbe-field">
                <span>{t('links.dialog.address')}</span>
                <input
                  ref={first}
                  value={draft.address}
                  inputMode="url"
                  autoComplete="off"
                  spellCheck={false}
                  onChange={(e) => set({ address: e.target.value })}
                  aria-invalid={!!error}
                  aria-describedby={describedBy('address', !!error)}
                />
              </label>
              <p id={`${id}-address`} className="wbe-muted wbe-small">
                {t('links.dialog.addressHint')}
              </p>
            </>
          ) : (
            <>
              <div className="wbe-link-place">
                <label className="wbe-field">
                  <span>{t('links.dialog.sheet')}</span>
                  <select ref={first} value={draft.place} onChange={(e) => set({ place: e.target.value })}>
                    <optgroup label={t('links.dialog.sheets')}>
                      {sheets.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </optgroup>
                    {names.length > 0 && (
                      <optgroup label={t('links.dialog.names')}>
                        {names.map((name) => (
                          <option key={name} value={NAME + name}>
                            {name}
                          </option>
                        ))}
                      </optgroup>
                    )}
                  </select>
                </label>
                <label className="wbe-field">
                  <span>{t('links.dialog.cell')}</span>
                  <input
                    ref={cellInput}
                    value={byName ? '' : draft.cell}
                    disabled={byName}
                    spellCheck={false}
                    autoComplete="off"
                    onChange={(e) => set({ cell: e.target.value })}
                    aria-invalid={!!error}
                    aria-describedby={describedBy('cell', !!error)}
                  />
                </label>
              </div>
              <p id={`${id}-cell`} className="wbe-muted wbe-small">
                {t('links.dialog.cellHint')}
              </p>
            </>
          )}
          {error && (
            <p id={`${id}-error`} className="wbe-link-error">
              {t(error)}
            </p>
          )}
          <label className="wbe-field">
            <span>{t('links.dialog.tooltip')}</span>
            <input value={draft.tooltip} onChange={(e) => set({ tooltip: e.target.value })} aria-describedby={`${id}-tip`} />
          </label>
          <p id={`${id}-tip`} className="wbe-muted wbe-small">
            {t('links.dialog.tooltipHint')}
          </p>
          <div className="wbe-modal-actions">
            {linked > 0 && (
              <button type="button" className="wbe-btn wbe-link-remove" onClick={remove}>
                {t('links.dialog.remove')}
              </button>
            )}
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t(editing ? 'links.dialog.save' : 'links.dialog.insert')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

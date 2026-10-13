import { useId, useRef, useState } from 'react';
import { displayFormula } from '@ruled/formula/formula';
import type { DefinedName } from '../core/types';
import type { Text } from '../i18n/message';
import { useT } from '../i18n/react';
import {
  defineName,
  deleteName,
  listedNames,
  nameError,
  nameHasError,
  nameValue,
  openDefineName,
  refersToError,
  refersToFile,
} from '../state/names';
import type { NameDialog as NameDialogState, WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';
import { Select } from './Select';

const refersToText = (n: DefinedName) => displayFormula(n.formula);

function openEditName(store: WorkbookStore, n: DefinedName): void {
  store.set('nameDialog', { from: { name: n.name, localSheet: n.localSheet }, name: n.name, localSheet: n.localSheet, refersTo: refersToText(n), comment: n.comment ?? '' });
}

/**
 * Defines a name or edits one, as Excel's New Name and Edit Name dialogs do:
 * its name, its scope (the workbook or one sheet, fixed once defined), a
 * comment, and what it refers to. Problems show under their field once OK
 * has been pressed, and the focus goes to the first.
 */
export function NameDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const req = store.nameDialog;
  const dialog = useRef<HTMLFormElement>(null);
  const nameInput = useRef<HTMLInputElement>(null);
  const refersInput = useRef<HTMLInputElement>(null);
  const id = useId();
  const [name, setName] = useState('');
  const [scope, setScope] = useState('');
  const [comment, setComment] = useState('');
  const [refersTo, setRefersTo] = useState('');
  const [tried, setTried] = useState(false);
  const [refused, setRefused] = useState<Text>();
  const opened = useRef<NameDialogState | undefined>(undefined);
  if (opened.current !== req) {
    opened.current = req;
    if (req) {
      setName(req.name);
      setScope(req.localSheet === undefined ? '' : String(req.localSheet));
      setComment(req.comment);
      setRefersTo(req.refersTo);
      setTried(false);
      setRefused(undefined);
    }
  }
  const close = () => store.set('nameDialog', undefined);
  useDialog(dialog, !!req, close, { modal: true, initial: nameInput });
  if (!req) return null;

  const editing = !!req.from;
  const localSheet = scope === '' ? undefined : Number(scope);
  const nameText = name.trim();
  const nameProblem = tried ? (refused ?? nameError(store, nameText, localSheet, req.from)) : undefined;
  const refersProblem = tried ? refersToError(store, refersTo) : undefined;
  const submit = async () => {
    setTried(true);
    setRefused(undefined);
    const problem = nameError(store, nameText, localSheet, req.from);
    if (problem || refersToError(store, refersTo)) {
      (problem ? nameInput : refersInput).current?.focus();
      return;
    }
    // What a name refers to is kept as the file has it unless it was changed here.
    const formula = editing && refersTo === req.refersTo ? store.names.find((n) => n.name === req.from!.name && n.localSheet === req.from!.localSheet)?.formula : undefined;
    const refusal = await defineName(store, { name: nameText, formula: formula ?? refersToFile(store, refersTo), localSheet, ...(comment.trim() && { comment: comment.trim() }) }, req.from);
    if (refusal) {
      setRefused(refusal);
      nameInput.current?.focus();
    } else {
      close();
    }
  };
  const sheets = store.sheets.flatMap((s, i) => (s.info.kind === 'worksheet' || i === localSheet ? [{ i, name: s.info.name }] : []));

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-name-form"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        noValidate
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          void submit();
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t(editing ? 'names.edit.title' : 'names.define.title')}</h2>
          <div className="wbe-field">
            <label htmlFor={`${id}-name`}>{t('names.field.name')}</label>
            <input
              id={`${id}-name`}
              ref={nameInput}
              value={name}
              spellCheck={false}
              autoComplete="off"
              maxLength={255}
              onChange={(e) => {
                setName(e.target.value);
                setRefused(undefined);
              }}
              aria-invalid={!!nameProblem}
              aria-describedby={nameProblem ? `${id}-name-error` : undefined}
            />
            {nameProblem && (
              <p id={`${id}-name-error`} className="wbe-field-error">
                {t.text(nameProblem)}
              </p>
            )}
          </div>
          <div className="wbe-field">
            <label htmlFor={`${id}-scope`}>{t('names.field.scope')}</label>
            <Select
              id={`${id}-scope`}
              value={scope}
              disabled={editing}
              onChange={(e) => setScope(e.target.value)}
              // Enter takes OK from here too, as from the other fields; Firefox doesn't submit from a list.
              onKeyDown={(e) => {
                if (e.key !== 'Enter' || e.altKey) return;
                e.preventDefault();
                e.currentTarget.form?.requestSubmit();
              }}
              aria-describedby={editing ? `${id}-scope-note` : undefined}
            >
              <option value="">{t('names.scope.workbook')}</option>
              {sheets.map((s) => (
                <option key={s.i} value={String(s.i)}>
                  {s.name}
                </option>
              ))}
            </Select>
            {editing && (
              <p id={`${id}-scope-note`} className="wbe-muted wbe-small">
                {t('names.scope.fixed')}
              </p>
            )}
          </div>
          <div className="wbe-field">
            <label htmlFor={`${id}-comment`}>{t('names.field.comment')}</label>
            <textarea id={`${id}-comment`} value={comment} rows={2} maxLength={255} onChange={(e) => setComment(e.target.value)} />
          </div>
          <div className="wbe-field">
            <label htmlFor={`${id}-refers`}>{t('names.field.refersTo')}</label>
            <input
              id={`${id}-refers`}
              ref={refersInput}
              className="wbe-mono"
              value={refersTo}
              spellCheck={false}
              autoComplete="off"
              onChange={(e) => setRefersTo(e.target.value)}
              aria-invalid={!!refersProblem}
              aria-describedby={refersProblem ? `${id}-refers-error` : undefined}
            />
            {refersProblem && (
              <p id={`${id}-refers-error`} className="wbe-field-error">
                {t.text(refersProblem)}
              </p>
            )}
          </div>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('names.define.submit')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

type Filter = 'all' | 'workbook' | 'sheet' | 'errors' | 'noErrors';

const keyOf = (n: DefinedName) => `${n.localSheet ?? ''}:${n.name}`;

/**
 * Excel's Name Manager: every name people define, with its value where it
 * can be read without calculating, what it refers to, its scope and comment,
 * to filter, and to add to, edit and delete. The names are a group of radio
 * buttons, so the arrow keys pick one; Enter edits it and Delete deletes it.
 */
export function NameManager({ store }: { store: WorkbookStore }) {
  const t = useT();
  const open = store.nameManagerOpen;
  const dialog = useRef<HTMLDivElement>(null);
  const id = useId();
  const [filter, setFilter] = useState<Filter>('all');
  const [picked, setPicked] = useState<string>();
  const close = () => store.set('nameManagerOpen', false);
  useDialog(dialog, open, close, { modal: true });
  if (!open) return null;

  const all = listedNames(store);
  const names = all.filter((n) => {
    if (filter === 'workbook') return n.localSheet === undefined;
    if (filter === 'sheet') return n.localSheet !== undefined;
    if (filter === 'errors') return nameHasError(store, n);
    if (filter === 'noErrors') return !nameHasError(store, n);
    return true;
  });
  const current = names.find((n) => keyOf(n) === picked) ?? names[0];
  const scopeName = (n: DefinedName) => (n.localSheet === undefined ? t('names.scope.workbook') : (store.sheets[n.localSheet]?.info.name ?? ''));
  const remove = async (n: DefinedName) => {
    if (!(await deleteName(store, n))) return;
    setPicked(undefined);
    // Its row has gone, and the focus with it: on to the name now picked, or the dialog.
    setTimeout(() => (dialog.current?.querySelector<HTMLElement>('input:checked') ?? dialog.current)?.focus());
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <div ref={dialog} className="wbe-modal wbe-names-manager" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1} onMouseDown={(e) => e.stopPropagation()}>
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('names.manager.title')}</h2>
          <div className="wbe-names-toolbar">
            <button type="button" className="wbe-btn" onClick={() => openDefineName(store)}>
              {t('names.manager.new')}
            </button>
            <button type="button" className="wbe-btn" disabled={!current} onClick={() => current && openEditName(store, current)}>
              {t('names.manager.edit')}
            </button>
            <button type="button" className="wbe-btn" disabled={!current} onClick={() => current && void remove(current)}>
              {t('names.manager.delete')}
            </button>
            <label className="wbe-names-filter">
              <span>{t('names.manager.filter')}</span>
              <Select value={filter} onChange={(e) => setFilter(e.target.value as Filter)}>
                {(['all', 'workbook', 'sheet', 'errors', 'noErrors'] as const).map((f) => (
                  <option key={f} value={f}>
                    {t(`names.filter.${f}`)}
                  </option>
                ))}
              </Select>
            </label>
          </div>
          {names.length ? (
            <div className="wbe-modal-table wbe-names-table">
              <table aria-label={t('names.manager.list')}>
                <thead>
                  <tr>
                    <th scope="col">{t('names.col.name')}</th>
                    <th scope="col">{t('names.col.value')}</th>
                    <th scope="col">{t('names.col.refersTo')}</th>
                    <th scope="col">{t('names.col.scope')}</th>
                    <th scope="col">{t('names.col.comment')}</th>
                  </tr>
                </thead>
                <tbody>
                  {names.map((n) => {
                    const on = n === current;
                    return (
                      <tr key={keyOf(n)} className={on ? 'wbe-picked' : ''} onClick={() => setPicked(keyOf(n))} onDoubleClick={() => openEditName(store, n)}>
                        <td>
                          <label className="wbe-names-pick">
                            <input
                              type="radio"
                              name={`${id}-pick`}
                              checked={on}
                              onChange={() => setPicked(keyOf(n))}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  openEditName(store, n);
                                } else if (e.key === 'Delete') {
                                  e.preventDefault();
                                  void remove(n);
                                }
                              }}
                            />
                            {n.name}
                          </label>
                        </td>
                        <td className="wbe-mono">{nameValue(store, n)}</td>
                        <td className="wbe-mono">{refersToText(n)}</td>
                        <td>{scopeName(n)}</td>
                        <td>{n.comment ?? ''}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="wbe-names-empty">{t(all.length ? 'names.manager.noMatch' : 'names.manager.empty')}</p>
          )}
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('names.manager.close')}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

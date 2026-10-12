import { useId, useRef, useState } from 'react';
import { useT } from '../i18n/react';
import { displayText } from '../state/format';
import { listedFormatCodes, usableFormat } from '../state/numberFormats';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';

/**
 * Excel's Custom number formats, from Format Cells: the active cell's format
 * code to change or type afresh, Excel's built-in codes and the workbook's own
 * to pick from, and the active cell as the code would show it. OK formats the
 * selection with the code.
 */
export function NumberFormatDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const open = store.numberFormatOpen;
  const id = useId();
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLInputElement>(null);
  const [code, setCode] = useState('');
  const [invalid, setInvalid] = useState(false);
  const opened = useRef(false);
  if (opened.current !== open) {
    opened.current = open;
    if (open) {
      setCode(store.activeStyle.numFmt || 'General');
      setInvalid(false);
    }
  }
  const close = () => store.set('numberFormatOpen', false);
  useDialog(dialog, open, close, { modal: true, initial });
  if (!open) return null;

  const codes = listedFormatCodes(store.styles.map((s) => s.numFmt), store.locale);
  const usable = usableFormat(code);
  const rec = store.view?.get(store.selection.row, store.selection.col);
  const sample = rec && usable ? displayText(rec, { ...store.activeStyle, numFmt: code }, store.date1904, false, store.locale) : undefined;
  const change = (next: string) => {
    setCode(next);
    setInvalid(false);
  };
  const submit = () => {
    if (!usable) {
      setInvalid(true);
      initial.current?.focus();
      return;
    }
    close();
    void store.formatCells({ numFmt: code });
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-number-format-dialog"
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
          <h2 id={`${id}-title`}>{t('format.numberFormat')}</h2>
          <div className="wbe-field">
            <span id={`${id}-sample-label`}>{t('numberFormat.sample')}</span>
            <output className="wbe-number-sample" aria-labelledby={`${id}-sample-label`} style={{ color: sample?.color }}>
              {sample?.text}
            </output>
          </div>
          <div className="wbe-field">
            <label htmlFor={`${id}-type`}>{t('numberFormat.type')}</label>
            <input
              id={`${id}-type`}
              ref={initial}
              value={code}
              autoComplete="off"
              spellCheck={false}
              onChange={(e) => change(e.target.value)}
              aria-invalid={invalid}
              aria-describedby={`${id}-hint${invalid ? ` ${id}-error` : ''}`}
            />
            <p id={`${id}-hint`} className="wbe-muted wbe-small">
              {t('numberFormat.hint')}
            </p>
            {invalid && (
              <p id={`${id}-error`} className="wbe-field-error">
                {t('numberFormat.invalid')}
              </p>
            )}
          </div>
          <select
            className="wbe-number-list"
            size={8}
            aria-label={t('numberFormat.list')}
            value={codes.includes(code) ? code : ''}
            onChange={(e) => change(e.target.value)}
            onDoubleClick={submit}
          >
            {codes.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('dialog.ok')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

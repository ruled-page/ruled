import { useId, useRef, useState } from 'react';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';

/**
 * Excel's Insert PivotTable dialog: the table or range to analyse, and
 * where the pivot table goes, on a new worksheet or at a cell of one.
 */
export function PivotDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const id = useId();
  const at = store.pivotDialog;
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLInputElement>(null);
  const [source, setSource] = useState('');
  const [existing, setExisting] = useState(false);
  const [location, setLocation] = useState('');
  const opened = useRef<typeof at>(undefined);
  if (opened.current !== at) {
    opened.current = at;
    if (at) {
      setSource(at.source);
      setLocation(at.location);
      setExisting(false);
    }
  }
  const close = () => store.set('pivotDialog', undefined);
  useDialog(dialog, !!at, close, { modal: true, initial });
  if (!at) return null;
  const ready = !!source.trim() && (!existing || !!location.trim());

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-pivot-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        noValidate
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          if (!ready) return;
          close();
          void store.insertPivot(source.trim(), existing ? location.trim() : undefined);
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('pivot.insert.heading')}</h2>
          <div className="wbe-field">
            <label htmlFor={`${id}-source`}>{t('pivot.insert.source')}</label>
            <input id={`${id}-source`} ref={initial} value={source} autoComplete="off" spellCheck={false} onChange={(e) => setSource(e.target.value)} aria-describedby={`${id}-source-hint`} />
            <p id={`${id}-source-hint`} className="wbe-muted wbe-small">
              {t('pivot.insert.sourceHint')}
            </p>
          </div>
          <fieldset className="wbe-choices">
            <legend>{t('pivot.insert.where')}</legend>
            <label className="wbe-choice">
              <input type="radio" name={`${id}-where`} checked={!existing} onChange={() => setExisting(false)} />
              {t('pivot.insert.newSheet')}
            </label>
            <label className="wbe-choice">
              <input type="radio" name={`${id}-where`} checked={existing} onChange={() => setExisting(true)} />
              {t('pivot.insert.existing')}
            </label>
          </fieldset>
          <div className="wbe-field">
            <label htmlFor={`${id}-location`}>{t('pivot.insert.location')}</label>
            <input id={`${id}-location`} value={location} disabled={!existing} autoComplete="off" spellCheck={false} onChange={(e) => setLocation(e.target.value)} />
          </div>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary" disabled={!ready}>
              {t('dialog.ok')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

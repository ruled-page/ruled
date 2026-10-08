import { useId, useRef } from 'react';
import type { Toast, WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { useDialog } from './useDialog';
import { useT } from '../i18n/react';

export function ConfirmDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const req = store.confirmRequest;
  const confirmBtn = useRef<HTMLButtonElement>(null);
  const cancelBtn = useRef<HTMLButtonElement>(null);
  const dialog = useRef<HTMLDivElement>(null);
  const id = useId();
  const danger = req?.danger !== false;
  // Something that can't be undone starts on the safe choice, so a stray Enter doesn't do it.
  useDialog(dialog, !!req, () => req?.resolve('cancel'), { modal: true, initial: danger ? cancelBtn : confirmBtn });
  if (!req) return null;
  return (
    <div className="wbe-modal-backdrop" onMouseDown={() => req.resolve('cancel')}>
      <div
        ref={dialog}
        className={`wbe-modal ${req.details ? 'wbe-modal-wide' : ''}`}
        role="alertdialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        aria-describedby={`${id}-body`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className={`wbe-modal-icon ${danger ? '' : 'wbe-calm'}`}>
          <Icon name={danger ? 'warn' : 'info'} size={20} />
        </div>
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t.text(req.title)}</h2>
          <p id={`${id}-body`}>{t.text(req.body)}</p>
          {req.details && (
            <div className="wbe-modal-table">
              <table>
                <thead>
                  <tr>
                    {req.details.header.map((h, i) => (
                      <th key={i}>{t.text(h)}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {req.details.rows.map((row, i) => (
                    <tr key={i}>
                      {row.map((cell, j) => (
                        <td key={j} className={j === 0 ? 'wbe-mono' : ''}>
                          {cell === '' ? <span className="wbe-muted">{t('dialog.empty')}</span> : cell}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {!!req.details.more && <p className="wbe-muted wbe-small">{t('dialog.more', { count: req.details.more })}</p>}
            </div>
          )}
          <div className="wbe-modal-actions">
            <button ref={cancelBtn} className="wbe-btn" onClick={() => req.resolve('cancel')}>
              {req.cancelLabel ? t.text(req.cancelLabel) : t('dialog.cancel')}
            </button>
            {req.alternateLabel && (
              <button className="wbe-btn" onClick={() => req.resolve('alternate')}>
                {t.text(req.alternateLabel)}
              </button>
            )}
            <button ref={confirmBtn} className={`wbe-btn ${danger ? 'wbe-danger' : 'wbe-primary'}`} onClick={() => req.resolve('confirm')}>
              {t.text(req.confirmLabel)}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/**
 * Messages that come and go. They are live regions that are always there,
 * news in a polite one and errors in an assertive one, each saying only a
 * message as it arrives. What an action did, which isn't shown, is said in a
 * third. Messages wait while the pointer or focus is on them.
 */
export function Toasts({ store }: { store: WorkbookStore }) {
  const t = useT();
  const toast = (item: Toast) => (
    <div key={item.id} className={`wbe-toast wbe-toast-${item.kind}`}>
      <Icon name={item.kind === 'info' ? 'check' : item.kind === 'warn' ? 'info' : 'warn'} />
      <span>{t.text(item.text)}</span>
      <button
        className="wbe-btn wbe-icon wbe-ghost"
        onClick={() => {
          store.dismissToast(item.id);
          store.holdToasts(false);
          // The button goes with its message; the focus goes back to the sheet.
          store.focusGrid();
        }}
        aria-label={t('toast.dismiss')}
      >
        <Icon name="close" size={14} />
      </button>
    </div>
  );
  return (
    <>
      <div
        className="wbe-toasts"
        onMouseEnter={() => store.holdToasts(true)}
        onMouseLeave={() => store.holdToasts(false)}
        onFocus={() => store.holdToasts(true)}
        onBlur={() => store.holdToasts(false)}
      >
        <div className="wbe-toast-group" role="status" aria-atomic="false">
          {store.toasts.filter((x) => x.kind !== 'error').map(toast)}
        </div>
        <div className="wbe-toast-group" role="alert" aria-atomic="false">
          {store.toasts.filter((x) => x.kind === 'error').map(toast)}
        </div>
      </div>
      <div className="wbe-sr-only" role="status" aria-atomic="false">
        {store.spoken.map((s) => (
          <p key={s.id}>{t.text(s.text)}</p>
        ))}
      </div>
    </>
  );
}

import { useId, useRef } from 'react';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { useDialog } from './useDialog';
import { useT } from '../i18n/react';

export function ConfirmDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const req = store.confirmRequest;
  const confirmBtn = useRef<HTMLButtonElement>(null);
  const dialog = useRef<HTMLDivElement>(null);
  const id = useId();
  useDialog(dialog, !!req, () => req?.resolve('cancel'), { modal: true, initial: confirmBtn });
  if (!req) return null;
  const danger = req.danger !== false;
  return (
    <div className="modal-backdrop" onMouseDown={() => req.resolve('cancel')}>
      <div
        ref={dialog}
        className={`modal ${req.details ? 'modal-wide' : ''}`}
        role="alertdialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        aria-describedby={`${id}-body`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className={`modal-icon ${danger ? '' : 'calm'}`}>
          <Icon name={danger ? 'warn' : 'info'} size={20} />
        </div>
        <div className="modal-content">
          <h2 id={`${id}-title`}>{t.text(req.title)}</h2>
          <p id={`${id}-body`}>{t.text(req.body)}</p>
          {req.details && (
            <div className="modal-table">
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
                        <td key={j} className={j === 0 ? 'mono' : ''}>
                          {cell === '' ? <span className="muted">{t('dialog.empty')}</span> : cell}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {!!req.details.more && <p className="muted small">{t('dialog.more', { count: req.details.more })}</p>}
            </div>
          )}
          <div className="modal-actions">
            <button className="btn" onClick={() => req.resolve('cancel')}>
              {req.cancelLabel ? t.text(req.cancelLabel) : t('dialog.cancel')}
            </button>
            {req.alternateLabel && (
              <button className="btn" onClick={() => req.resolve('alternate')}>
                {t.text(req.alternateLabel)}
              </button>
            )}
            <button ref={confirmBtn} className={`btn ${danger ? 'danger' : 'primary'}`} onClick={() => req.resolve('confirm')}>
              {t.text(req.confirmLabel)}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export function Toasts({ store }: { store: WorkbookStore }) {
  const t = useT();
  return (
    <div className="toasts" role="status">
      {store.toasts.map((toast) => (
        // Errors interrupt; the rest wait for a pause.
        <div key={toast.id} className={`toast toast-${toast.kind}`} role={toast.kind === 'error' ? 'alert' : undefined}>
          <Icon name={toast.kind === 'info' ? 'check' : toast.kind === 'warn' ? 'info' : 'warn'} />
          <span>{t.text(toast.text)}</span>
          <button className="btn icon ghost" onClick={() => store.dismissToast(toast.id)} aria-label={t('toast.dismiss')}>
            <Icon name="close" size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}

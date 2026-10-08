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
            <button className="wbe-btn" onClick={() => req.resolve('cancel')}>
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

export function Toasts({ store }: { store: WorkbookStore }) {
  const t = useT();
  return (
    <div className="wbe-toasts" role="status">
      {store.toasts.map((toast) => (
        // Errors interrupt; the rest wait for a pause.
        <div key={toast.id} className={`wbe-toast wbe-toast-${toast.kind}`} role={toast.kind === 'error' ? 'alert' : undefined}>
          <Icon name={toast.kind === 'info' ? 'check' : toast.kind === 'warn' ? 'info' : 'warn'} />
          <span>{t.text(toast.text)}</span>
          <button className="wbe-btn wbe-icon wbe-ghost" onClick={() => store.dismissToast(toast.id)} aria-label={t('toast.dismiss')}>
            <Icon name="close" size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}

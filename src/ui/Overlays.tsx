import { useEffect, useRef } from 'react';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';

export function ConfirmDialog({ store }: { store: WorkbookStore }) {
  const req = store.confirmRequest;
  const confirmBtn = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (!req) return;
    confirmBtn.current?.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') req.resolve('cancel');
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [req]);
  if (!req) return null;
  const danger = req.danger !== false;
  return (
    <div className="modal-backdrop" onMouseDown={() => req.resolve('cancel')}>
      <div
        className={`modal ${req.details ? 'modal-wide' : ''}`}
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="confirm-title"
        aria-describedby="confirm-body"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className={`modal-icon ${danger ? '' : 'calm'}`}>
          <Icon name={danger ? 'warn' : 'info'} size={20} />
        </div>
        <div className="modal-content">
          <h2 id="confirm-title">{req.title}</h2>
          <p id="confirm-body">{req.body}</p>
          {req.details && (
            <div className="modal-table">
              <table>
                <thead>
                  <tr>
                    {req.details.header.map((h) => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {req.details.rows.map((row, i) => (
                    <tr key={i}>
                      {row.map((cell, j) => (
                        <td key={j} className={j === 0 ? 'mono' : ''}>
                          {cell === '' ? <span className="muted">(empty)</span> : cell}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {!!req.details.more && <p className="muted small">…and {req.details.more.toLocaleString()} more.</p>}
            </div>
          )}
          <div className="modal-actions">
            <button className="btn" onClick={() => req.resolve('cancel')}>
              {req.cancelLabel ?? 'Cancel'}
            </button>
            {req.alternateLabel && (
              <button className="btn" onClick={() => req.resolve('alternate')}>
                {req.alternateLabel}
              </button>
            )}
            <button ref={confirmBtn} className={`btn ${danger ? 'danger' : 'primary'}`} onClick={() => req.resolve('confirm')}>
              {req.confirmLabel}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export function Toasts({ store }: { store: WorkbookStore }) {
  return (
    <div className="toasts" aria-live="polite">
      {store.toasts.map((t) => (
        <div key={t.id} className={`toast toast-${t.kind}`}>
          <Icon name={t.kind === 'info' ? 'check' : t.kind === 'warn' ? 'info' : 'warn'} />
          <span>{t.text}</span>
          <button className="btn icon ghost" onClick={() => store.dismissToast(t.id)} aria-label="Dismiss">
            <Icon name="close" size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}

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
      if (e.key === 'Escape') req.resolve(false);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [req]);
  if (!req) return null;
  return (
    <div className="modal-backdrop" onMouseDown={() => req.resolve(false)}>
      <div className="modal" role="alertdialog" aria-modal="true" aria-labelledby="confirm-title" onMouseDown={(e) => e.stopPropagation()}>
        <div className="modal-icon">
          <Icon name="warn" size={20} />
        </div>
        <div>
          <h2 id="confirm-title">{req.title}</h2>
          <p>{req.body}</p>
          <div className="modal-actions">
            <button className="btn" onClick={() => req.resolve(false)}>
              Cancel
            </button>
            <button ref={confirmBtn} className="btn danger" onClick={() => req.resolve(true)}>
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

import { useEffect, useId, useRef, useState } from 'react';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { useDialog } from './useDialog';

/** Asks for a password-protected workbook's password, as Excel does on opening it; again, saying so, after a wrong one. */
export function PasswordDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const req = store.passwordRequest;
  const dialog = useRef<HTMLFormElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const id = useId();
  const [password, setPassword] = useState('');
  // Each time it asks, it starts empty, as Excel's does after a wrong password.
  useEffect(() => setPassword(''), [req]);
  useDialog(dialog, !!req, () => req?.resolve(undefined), { modal: true, initial: input });
  if (!req) return null;
  return (
    <div className="wbe-modal-backdrop" onMouseDown={() => req.resolve(undefined)}>
      <form
        ref={dialog}
        className="wbe-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        aria-describedby={`${id}-body`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          if (password) req.resolve(password);
        }}
      >
        <div className="wbe-modal-icon wbe-calm">
          <Icon name="info" size={20} />
        </div>
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('password.title')}</h2>
          <p id={`${id}-body`}>{t('password.body', { fileName: req.fileName })}</p>
          <label className="wbe-field">
            <span>{t('password.label')}</span>
            <input
              ref={input}
              type="password"
              autoComplete="off"
              value={password}
              aria-invalid={req.wrong || undefined}
              aria-describedby={req.wrong ? `${id}-wrong` : undefined}
              onChange={(e) => setPassword(e.target.value)}
            />
          </label>
          {req.wrong && (
            <p id={`${id}-wrong`} className="wbe-field-error" role="alert">
              {t('password.wrong')}
            </p>
          )}
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={() => req.resolve(undefined)}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary" disabled={!password}>
              {t('password.open')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

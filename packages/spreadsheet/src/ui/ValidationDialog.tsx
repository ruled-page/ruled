import { useId, useRef, useState } from 'react';
import { listSeparator } from '../core/locale';
import type { RuleValidation } from '../core/types';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { draftFor, ruleFromDraft, type DraftProblem, type ValidationAllow, type ValidationDraft } from '../state/validationDraft';
import { Select } from './Select';
import { Tabs } from './Tabs';
import { useDialog } from './useDialog';

const ALLOW: ValidationAllow[] = ['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'];
const OPERATORS: RuleValidation['operator'][] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
type Tab = 'settings' | 'input' | 'error';

/** What Excel calls a rule's boxes, for the kind of value and the comparison: Minimum and Maximum, Start date, Length… */
function boxNames(allow: ValidationAllow, operator: RuleValidation['operator']): [MessageKey, MessageKey] {
  const [low, high, one]: [MessageKey, MessageKey, MessageKey] =
    allow === 'date'
      ? ['validation.startDate', 'validation.endDate', 'validation.date']
      : allow === 'time'
        ? ['validation.startTime', 'validation.endTime', 'validation.time']
        : allow === 'textLength'
          ? ['validation.minimum', 'validation.maximum', 'validation.length']
          : ['validation.minimum', 'validation.maximum', 'validation.value'];
  if (operator === 'between' || operator === 'notBetween') return [low, high];
  if (operator === 'greaterThan' || operator === 'greaterThanOrEqual') return [low, high];
  if (operator === 'lessThan' || operator === 'lessThanOrEqual') return [high, high];
  return [one, one];
}

/**
 * Excel's Data Validation dialog for the selection: what its cells allow
 * (any value, whole numbers, decimals, a list, dates, times, a text length or
 * a formula), the message shown when one is selected, and the alert when
 * something else is typed. Clear All takes the cells' rule away.
 */
export function ValidationDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const at = store.validationDialog;
  const id = useId();
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLSelectElement>(null);
  const [draft, setDraft] = useState<ValidationDraft | null>(null);
  const [tab, setTab] = useState<Tab>('settings');
  const [problem, setProblem] = useState<DraftProblem>();
  const opened = useRef<typeof at>(undefined);
  if (opened.current !== at) {
    opened.current = at;
    if (at) {
      setDraft(draftFor(at.rule, store.validationPlace(at.rects), store.styles[0]));
      setTab('settings');
      setProblem(undefined);
    }
  }
  const close = () => store.set('validationDialog', undefined);
  useDialog(dialog, !!at, close, { modal: true, initial });
  if (!at || !draft) return null;

  const set = (patch: Partial<ValidationDraft>) => {
    setDraft({ ...draft, ...patch });
    setProblem(undefined);
  };
  const submit = () => {
    const made = ruleFromDraft(draft, store.validationPlace(at.rects), store.listContext());
    if ('problem' in made) {
      setTab('settings');
      setProblem(made.problem);
      return;
    }
    close();
    void store.setValidation(at.rects, made.rule);
  };
  const clearAll = () => {
    close();
    void store.setValidation(at.rects, undefined);
  };

  const numeric = draft.allow !== 'any' && draft.allow !== 'list' && draft.allow !== 'custom';
  const two = draft.operator === 'between' || draft.operator === 'notBetween';
  const [firstName, secondName] = boxNames(draft.allow, draft.operator);
  const field = (name: 'first' | 'second', label: MessageKey, hint?: string) => {
    const error = problem?.field === name ? problem.message : undefined;
    return (
      <div className="wbe-field">
        <label htmlFor={`${id}-${name}`}>{t(label)}</label>
        <input
          id={`${id}-${name}`}
          value={draft[name]}
          autoComplete="off"
          spellCheck={false}
          onChange={(e) => set({ [name]: e.target.value })}
          aria-invalid={!!error}
          aria-describedby={[hint && `${id}-${name}-hint`, error && `${id}-${name}-error`].filter(Boolean).join(' ') || undefined}
        />
        {hint && (
          <p id={`${id}-${name}-hint`} className="wbe-muted wbe-small">
            {hint}
          </p>
        )}
        {error && (
          <p id={`${id}-${name}-error`} className="wbe-field-error">
            {t(error)}
          </p>
        )}
      </div>
    );
  };
  const check = (key: 'allowBlank' | 'inCellDropdown' | 'showPrompt' | 'showError', label: MessageKey) => (
    <label className="wbe-validation-check">
      <input type="checkbox" checked={draft[key]} onChange={(e) => set({ [key]: e.target.checked })} />
      {t(label)}
    </label>
  );
  const text = (key: 'promptTitle' | 'prompt' | 'errorTitle' | 'error', label: MessageKey, long?: boolean) => (
    <div className="wbe-field">
      <label htmlFor={`${id}-${key}`}>{t(label)}</label>
      {long ? (
        <textarea id={`${id}-${key}`} rows={4} value={draft[key]} onChange={(e) => set({ [key]: e.target.value })} />
      ) : (
        <input id={`${id}-${key}`} value={draft[key]} autoComplete="off" onChange={(e) => set({ [key]: e.target.value })} />
      )}
    </div>
  );

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-validation-dialog"
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
          <h2 id={`${id}-title`}>{t('validation.title')}</h2>
          <Tabs<Tab>
            tabs={[
              { id: 'settings', label: t('validation.tab.settings') },
              { id: 'input', label: t('validation.tab.input') },
              { id: 'error', label: t('validation.tab.error') },
            ]}
            value={tab}
            onChange={setTab}
          >
            <div className="wbe-validation-panel">
              {tab === 'settings' && (
                <>
                  <div className="wbe-validation-row">
                    <div className="wbe-field">
                      <label htmlFor={`${id}-allow`}>{t('validation.allow')}</label>
                      <Select id={`${id}-allow`} ref={initial} value={draft.allow} onChange={(e) => set({ allow: e.target.value as ValidationAllow })}>
                        {ALLOW.map((a) => (
                          <option key={a} value={a}>
                            {t(`validation.allow.${a}`)}
                          </option>
                        ))}
                      </Select>
                    </div>
                    {numeric && (
                      <div className="wbe-field">
                        <label htmlFor={`${id}-operator`}>{t('validation.data')}</label>
                        <Select id={`${id}-operator`} value={draft.operator} onChange={(e) => set({ operator: e.target.value as RuleValidation['operator'] })}>
                          {OPERATORS.map((op) => (
                            <option key={op} value={op}>
                              {t(`validation.op.${op}`)}
                            </option>
                          ))}
                        </Select>
                      </div>
                    )}
                  </div>
                  {draft.allow !== 'any' && check('allowBlank', 'validation.ignoreBlank')}
                  {draft.allow === 'list' && (
                    <>
                      {check('inCellDropdown', 'validation.dropdown')}
                      {field('first', 'validation.source', t('validation.sourceHint', { separator: listSeparator(store.locale) }))}
                    </>
                  )}
                  {draft.allow === 'custom' && field('first', 'validation.formula')}
                  {numeric && field('first', firstName)}
                  {numeric && two && field('second', secondName)}
                </>
              )}
              {tab === 'input' && (
                <>
                  {check('showPrompt', 'validation.showPrompt')}
                  {text('promptTitle', 'validation.promptTitle')}
                  {text('prompt', 'validation.prompt', true)}
                </>
              )}
              {tab === 'error' && (
                <>
                  {check('showError', 'validation.showError')}
                  <div className="wbe-field">
                    <label htmlFor={`${id}-style`}>{t('validation.style')}</label>
                    <Select id={`${id}-style`} value={draft.errorStyle} onChange={(e) => set({ errorStyle: e.target.value as ValidationDraft['errorStyle'] })}>
                      {(['stop', 'warning', 'information'] as const).map((s) => (
                        <option key={s} value={s}>
                          {t(`validation.style.${s}`)}
                        </option>
                      ))}
                    </Select>
                  </div>
                  {text('errorTitle', 'validation.errorTitle')}
                  {text('error', 'validation.error', true)}
                </>
              )}
            </div>
          </Tabs>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn wbe-validation-clear" onClick={clearAll}>
              {t('validation.clearAll')}
            </button>
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

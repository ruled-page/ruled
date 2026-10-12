import { useId, useMemo, useRef, useState } from 'react';
import { colToLetters } from '@sitepen/formula-engine/refs';
import { VT, type FilterCondition } from '../core/types';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { RULE_OPS, ruleFor, ruleOp, typedRuleValue, type RuleOp } from '../state/filterConditions';
import { editText } from '../state/format';
import type { WorkbookStore } from '../state/store';
import { Select } from './Select';
import { useDialog } from './useDialog';

/** How a date column names its comparisons, as Excel's dialog does. */
const DATE_NAMES: Partial<Record<RuleOp, MessageKey>> = {
  greaterThan: 'filter.rule.after',
  greaterThanOrEqual: 'filter.rule.afterOrEqual',
  lessThan: 'filter.rule.before',
  lessThanOrEqual: 'filter.rule.beforeOrEqual',
};

interface Rule {
  op: RuleOp | '';
  text: string;
}

/**
 * Excel's Custom AutoFilter dialog, for one or two rules joined by And or
 * Or, and its Top 10 AutoFilter dialog, for a column of a filter. Numbers
 * and dates are typed as the locale writes them.
 */
export function FilterConditionDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const at = store.conditionDialog;
  const id = useId();
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLSelectElement>(null);
  const choices = useMemo(() => (at ? store.filterChoices(at.col, at.rect) : undefined), [store, at]);
  const [rules, setRules] = useState<[Rule, Rule]>([
    { op: 'equal', text: '' },
    { op: '', text: '' },
  ]);
  const [and, setAnd] = useState(true);
  const [top, setTop] = useState({ top: true, count: '10', percent: false });
  const [tried, setTried] = useState(false);
  const opened = useRef<typeof at>(undefined);
  if (opened.current !== at) {
    opened.current = at;
    if (at && choices) {
      const c = choices.condition;
      // A rule's value as it was typed: numbers and dates as the locale writes them.
      const shown = (value: string) => {
        if (choices.kind === 'text' || value.trim() === '' || !Number.isFinite(Number(value))) return value;
        const style = choices.kind === 'date' ? choices.style : 0;
        return editText({ t: VT.Number, n: Number(value), style, flags: 0 }, store.styles[style] ?? store.styles[0], store.date1904, store.locale);
      };
      if (c?.kind === 'custom') {
        const [a, b] = c.rules.map(ruleOp);
        setRules([
          { op: a.op, text: shown(a.value) },
          b ? { op: b.op, text: shown(b.value) } : { op: '', text: '' },
        ]);
        setAnd(c.and);
      } else {
        const p = at.preset;
        setRules([
          { op: p?.first ?? 'equal', text: '' },
          { op: p?.second ?? '', text: '' },
        ]);
        setAnd(p?.and ?? true);
      }
      setTop(c?.kind === 'top10' ? { top: c.top, count: String(c.count), percent: c.percent } : { top: true, count: '10', percent: false });
      setTried(false);
    }
  }
  const close = () => store.set('conditionDialog', undefined);
  useDialog(dialog, !!(at && choices), close, { modal: true, initial });
  if (!at || !choices) return null;

  const view = store.view;
  const heading = (view && store.shownText(view, choices.rect.r1, at.col)) || colToLetters(at.col);
  const kind = choices.kind;
  const ruleName = (i: number) => t(i === 0 ? 'filter.custom.first' : 'filter.custom.second');
  // Each rule as the file keeps it, or why it can't be.
  const read = rules.map((r) => (r.op === '' ? null : typedRuleValue(r.text, r.op, kind, store.locale, store.date1904)));
  const needs: MessageKey = kind === 'date' ? 'filter.custom.needDate' : 'filter.custom.needNumber';
  const count = Number(top.count);
  const countOk = Number.isInteger(count) && count >= 1 && count <= 500;
  const invalid = at.mode === 'top10' ? !countOk : read.some((v) => v === undefined);

  const submit = () => {
    if (invalid) return setTried(true);
    let condition: FilterCondition;
    if (at.mode === 'top10') condition = { kind: 'top10', top: top.top, percent: top.percent, count };
    else {
      const made = rules.flatMap((r, i) => (r.op === '' ? [] : [ruleFor(r.op, read[i]!)]));
      condition = { kind: 'custom', and, rules: made };
    }
    close();
    void store.setFilterCondition(choices.rect, at.col, condition);
  };

  const setRule = (i: number, patch: Partial<Rule>) => setRules(rules.map((r, j) => (j === i ? { ...r, ...patch } : r)) as [Rule, Rule]);
  const listId = `${id}-values`;
  const ruleRow = (i: number) => {
    const r = rules[i];
    const error = tried && read[i] === undefined;
    return (
      <div className="wbe-condition-rule" key={i}>
        <Select ref={i === 0 ? initial : undefined} aria-label={t('filter.custom.operator', { rule: ruleName(i) })} value={r.op} onChange={(e) => setRule(i, { op: e.target.value as RuleOp | '' })}>
          {i > 0 && <option value="">{t('filter.custom.none')}</option>}
          {RULE_OPS.map((op) => (
            <option key={op} value={op}>
              {t((kind === 'date' && DATE_NAMES[op]) || (`filter.rule.${op}` as MessageKey))}
            </option>
          ))}
        </Select>
        <input
          className="wbe-condition-value"
          aria-label={t('filter.custom.value', { rule: ruleName(i) })}
          value={r.text}
          list={listId}
          disabled={r.op === ''}
          autoComplete="off"
          onChange={(e) => setRule(i, { text: e.target.value })}
          aria-invalid={error}
          aria-describedby={error ? `${id}-error-${i}` : undefined}
        />
        {error && (
          <p id={`${id}-error-${i}`} className="wbe-field-error">
            {t(needs)}
          </p>
        )}
      </div>
    );
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-condition-dialog"
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
          <h2 id={`${id}-title`}>{t(at.mode === 'top10' ? 'filter.top10.title' : 'filter.custom.title')}</h2>
          {at.mode === 'top10' ? (
            <div className="wbe-condition-top10">
              <span>{t('filter.top10.show')}</span>
              <Select ref={initial} aria-label={t('filter.top10.show')} value={top.top ? 'top' : 'bottom'} onChange={(e) => setTop({ ...top, top: e.target.value === 'top' })}>
                <option value="top">{t('filter.top10.top')}</option>
                <option value="bottom">{t('filter.top10.bottom')}</option>
              </Select>
              <input
                className="wbe-condition-count"
                inputMode="numeric"
                aria-label={t('filter.top10.count')}
                value={top.count}
                onChange={(e) => setTop({ ...top, count: e.target.value })}
                aria-invalid={tried && !countOk}
                aria-describedby={tried && !countOk ? `${id}-count-error` : undefined}
              />
              <Select aria-label={t('filter.top10.items')} value={top.percent ? 'percent' : 'items'} onChange={(e) => setTop({ ...top, percent: e.target.value === 'percent' })}>
                <option value="items">{t('filter.top10.items')}</option>
                <option value="percent">{t('filter.top10.percent')}</option>
              </Select>
              {tried && !countOk && (
                <p id={`${id}-count-error`} className="wbe-field-error">
                  {t('filter.top10.invalid')}
                </p>
              )}
            </div>
          ) : (
            <>
              <p>{t('filter.custom.showRows', { heading })}</p>
              {ruleRow(0)}
              <div className="wbe-condition-join" role="radiogroup" aria-label={`${t('filter.custom.and')} / ${t('filter.custom.or')}`}>
                {[true, false].map((both) => (
                  <label key={String(both)}>
                    <input type="radio" name={`${id}-join`} checked={and === both} onChange={() => setAnd(both)} />
                    {t(both ? 'filter.custom.and' : 'filter.custom.or')}
                  </label>
                ))}
              </div>
              {ruleRow(1)}
              <p className="wbe-muted wbe-small">{t('filter.custom.hint')}</p>
              <datalist id={listId}>
                {choices.values
                  .filter((v) => v.text !== '')
                  .slice(0, 200)
                  .map((v) => (
                    <option key={v.text} value={v.text}>
                      {v.text}
                    </option>
                  ))}
              </datalist>
            </>
          )}
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

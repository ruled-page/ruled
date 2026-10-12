import { useId, useRef, useState, type CSSProperties } from 'react';
import { parseRange, rangeToA1 } from '@sitepen/formula-engine/refs';
import type { CfRule, DiffStyle, Rect } from '../core/types';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { CF_GALLERY, CF_PRESETS, CF_VALUES, describeCfRule, galleryRule, quickOf, quickRule } from '../state/cfRules';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { Select } from './Select';
import { useDialog } from './useDialog';

const sameStyle = (a: DiffStyle | undefined, b: DiffStyle) => JSON.stringify(a ?? {}) === JSON.stringify(b);

/** How a format looks, for the dialogs' samples. */
function sample(style: DiffStyle | undefined): CSSProperties {
  const b = style?.border?.top;
  return {
    background: style?.fill,
    color: style?.color,
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? 'italic' : undefined,
    textDecoration: [style?.underline && 'underline', style?.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
    outline: b ? `1px solid ${b.color ?? 'currentColor'}` : undefined,
  };
}

/**
 * Excel's quick conditional formatting dialogs: Greater Than, Between, Text
 * That Contains, Top 10 Items, Above Average, New Formatting Rule and the
 * rest, each a value or two (or a formula) and one of its formats. Opened
 * from the Rules Manager, it changes the rule there instead.
 */
export function CfDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const at = store.cfDialog;
  const id = useId();
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLInputElement>(null);
  const [values, setValues] = useState<string[]>([]);
  const [preset, setPreset] = useState('0');
  const [problem, setProblem] = useState<{ message: MessageKey; box: number }>();
  const opened = useRef<typeof at>(undefined);
  if (opened.current !== at) {
    opened.current = at;
    if (at) {
      const ranked = at.kind === 'top' || at.kind === 'topPercent' || at.kind === 'bottom' || at.kind === 'bottomPercent';
      setValues(at.values ?? (ranked ? ['10'] : ['', '']));
      // A rule's own format, when it isn't one of Excel's, is offered as it is.
      const own = at.style && CF_PRESETS.findIndex((p) => sameStyle(at.style, p.style));
      setPreset(own === undefined || own < 0 ? (at.style ? 'own' : '0') : String(own));
      setProblem(undefined);
    }
  }
  const close = () => store.set('cfDialog', undefined);
  useDialog(dialog, !!at, close, { modal: true, initial });
  if (!at) return null;

  const style = preset === 'own' && at.style ? at.style : CF_PRESETS[Number(preset)].style;
  const submit = () => {
    const made = quickRule(at.kind, values, style, store.validationPlace(at.rects));
    if ('problem' in made) return setProblem(made.problem);
    close();
    const manager = store.cfManager;
    if (at.edit !== undefined && manager) {
      const was = manager.rules[at.edit];
      // The changed rule keeps its place, its cells and Stop If True; a new format gets a place of its own.
      const rule: CfRule = { ...made.rule, rects: was.rects, stopIfTrue: was.stopIfTrue, ...(sameStyle(was.style, style) && was.dxfId !== undefined && { dxfId: was.dxfId }) };
      store.set('cfManager', { rules: manager.rules.map((r, i) => (i === at.edit ? rule : r)) });
    } else {
      void store.addConditionalFormat(made.rule);
    }
  };
  const count = CF_VALUES[at.kind];
  const box = (i: number) => {
    const error = problem?.box === i ? problem.message : undefined;
    return (
      <input
        key={i}
        ref={i === 0 ? initial : undefined}
        className="wbe-cf-value"
        aria-label={t(i === 0 ? 'cf.value' : 'cf.value2')}
        value={values[i] ?? ''}
        autoComplete="off"
        spellCheck={false}
        onChange={(e) => {
          setValues(values.map((v, j) => (j === i ? e.target.value : v)).concat(values.length <= i ? [e.target.value] : []));
          setProblem(undefined);
        }}
        aria-invalid={!!error}
        aria-describedby={error ? `${id}-problem` : undefined}
      />
    );
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-cf-dialog"
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
          <h2 id={`${id}-title`}>{t(`cf.title.${at.kind}`)}</h2>
          <p>{t(`cf.prompt.${at.kind}`)}</p>
          <div className="wbe-cf-row">
            {count > 0 && box(0)}
            {count > 1 && box(1)}
            <span>{t('cf.with')}</span>
            <Select aria-label={t('cf.with')} value={preset} onChange={(e) => setPreset(e.target.value)}>
              {CF_PRESETS.map((p, i) => (
                <option key={p.key} value={String(i)}>
                  {t(p.key)}
                </option>
              ))}
              {at.style && !CF_PRESETS.some((p) => sameStyle(at.style, p.style)) && <option value="own">{t('cf.manager.format')}</option>}
            </Select>
            <span className="wbe-cf-sample" style={sample(style)} aria-hidden="true">
              {t('cf.manager.preview')}
            </span>
          </div>
          {problem && (
            <p id={`${id}-problem`} className="wbe-field-error" role="alert">
              {t(problem.message)}
            </p>
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

/** A rule's ranges from what's typed: "A1:B10 D2", commas or spaces between them, dollar signs or not. */
function readRanges(text: string): Rect[] | undefined {
  const parts = text.replace(/^=/, '').replace(/\$/g, '').split(/[\s,;]+/).filter(Boolean);
  const rects = parts.map((p) => parseRange(p));
  return rects.length && rects.every(Boolean) ? (rects as Rect[]) : undefined;
}

/**
 * Excel's Conditional Formatting Rules Manager: the sheet's rules (or those
 * on the selection) in the order they apply, each with its format and its
 * cells, to edit, delete, reorder, re-range or stop after. Changes are made
 * together on OK, as one step.
 */
export function CfManager({ store }: { store: WorkbookStore }) {
  const t = useT();
  const m = store.cfManager;
  const id = useId();
  const dialog = useRef<HTMLDivElement>(null);
  const [scope, setScope] = useState<'sheet' | 'selection'>('sheet');
  const [picked, setPicked] = useState(0);
  const [ranges, setRanges] = useState<Record<number, string>>({});
  const opened = useRef<typeof m>(undefined);
  const editing = store.cfDialog?.edit !== undefined;
  if (!!opened.current !== !!m) {
    opened.current = m;
    setScope('sheet');
    setPicked(0);
    setRanges({});
  }
  const close = () => store.set('cfManager', undefined);
  useDialog(dialog, !!m && !editing, close, { modal: true });
  if (!m || editing) return null;

  const sel = store.selection.rects;
  const touches = (r: CfRule) => r.rects.some((a) => sel.some((s) => a.r1 <= s.r2 && a.r2 >= s.r1 && a.c1 <= s.c2 && a.c2 >= s.c1));
  const shown = m.rules.map((rule, index) => ({ rule, index })).filter(({ rule }) => scope === 'sheet' || touches(rule));
  const current = shown.find((s) => s.index === picked) ?? shown[0];
  const setRules = (rules: CfRule[]) => store.set('cfManager', { rules });
  const move = (by: number) => {
    if (!current) return;
    const to = current.index + by;
    const rules = [...m.rules];
    [rules[current.index], rules[to]] = [rules[to], rules[current.index]];
    setRules(rules);
    setPicked(to);
    // Ranges typed but not yet applied go with their rules.
    const next: Record<number, string> = { ...ranges };
    delete next[current.index];
    delete next[to];
    if (ranges[current.index] !== undefined) next[to] = ranges[current.index];
    if (ranges[to] !== undefined) next[current.index] = ranges[to];
    setRanges(next);
  };
  const bad = Object.entries(ranges).some(([, text]) => !readRanges(text));
  const edit = () => {
    if (!current) return;
    const quick = quickOf(current.rule, store.validationPlace(current.rule.rects));
    if (quick) store.set('cfDialog', { ...quick, rects: current.rule.rects, style: current.rule.style, edit: current.index });
  };
  const apply = () => {
    if (bad) return;
    const rules = m.rules.map((r, i) => (ranges[i] !== undefined ? { ...r, rects: readRanges(ranges[i])! } : r));
    close();
    if (JSON.stringify(rules) !== JSON.stringify(store.view?.info.conditionalFormats)) void store.setConditionalFormats(rules);
  };
  const look = (r: CfRule): CSSProperties =>
    r.colorScale
      ? { background: `linear-gradient(to right, ${r.colorScale.colors.join(', ')})` }
      : r.dataBar
        ? { background: `linear-gradient(to right, ${r.dataBar.color} 60%, transparent 60%)` }
        : sample(r.style);

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <div
        ref={dialog}
        className="wbe-modal wbe-modal-wide wbe-cf-manager"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('cf.manager.title')}</h2>
          <div className="wbe-field wbe-cf-scope">
            <label htmlFor={`${id}-scope`}>{t('cf.manager.show')}</label>
            <Select id={`${id}-scope`} value={scope} onChange={(e) => setScope(e.target.value as 'sheet' | 'selection')}>
              <option value="selection">{t('cf.manager.selection')}</option>
              <option value="sheet">{t('cf.manager.sheet')}</option>
            </Select>
          </div>
          <div className="wbe-cf-tools">
            <button type="button" className="wbe-btn" disabled={!current || current.rule.x14Only || !quickOf(current.rule, store.validationPlace(current.rule.rects))} onClick={edit}>
              {t('cf.manager.edit')}
            </button>
            <button
              type="button"
              className="wbe-btn"
              disabled={!current}
              onClick={() => {
                if (!current) return;
                setRules(m.rules.filter((_, i) => i !== current.index));
                setRanges({});
              }}
            >
              <Icon name="trash" />
              {t('cf.manager.delete')}
            </button>
            <button type="button" className="wbe-btn wbe-icon" disabled={!current || current.index === 0} onClick={() => move(-1)} aria-label={t('cf.manager.moveUp')}>
              <Icon name="chevronUp" />
            </button>
            <button type="button" className="wbe-btn wbe-icon" disabled={!current || current.index === m.rules.length - 1} onClick={() => move(1)} aria-label={t('cf.manager.moveDown')}>
              <Icon name="chevronDown" />
            </button>
          </div>
          {shown.length ? (
            <table className="wbe-cf-rules">
              <thead>
                <tr>
                  <th scope="col">{t('cf.manager.rule')}</th>
                  <th scope="col">{t('cf.manager.format')}</th>
                  <th scope="col">{t('cf.manager.appliesTo')}</th>
                  <th scope="col">{t('cf.manager.stopIfTrue')}</th>
                </tr>
              </thead>
              <tbody>
                {shown.map(({ rule, index }) => {
                  const text = ranges[index] ?? rule.rects.map(rangeToA1).join(' ');
                  const wrong = ranges[index] !== undefined && !readRanges(ranges[index]);
                  const name = t.text(describeCfRule(rule));
                  return (
                    <tr key={index} className={current?.index === index ? 'wbe-active' : undefined} aria-selected={current?.index === index} onClick={() => setPicked(index)}>
                      <td>
                        <label>
                          <input type="radio" name={`${id}-rule`} checked={current?.index === index} onChange={() => setPicked(index)} />
                          {name}
                        </label>
                        {rule.x14Only && <span className="wbe-muted wbe-small">{t('cf.manager.kept')}</span>}
                      </td>
                      <td>
                        <span className="wbe-cf-sample" style={look(rule)}>
                          {rule.colorScale || rule.dataBar || rule.iconSet ? ' ' : t('cf.manager.preview')}
                        </span>
                      </td>
                      <td>
                        <input
                          className="wbe-cf-ranges"
                          value={text}
                          aria-label={`${t('cf.manager.appliesTo')}: ${name}`}
                          aria-invalid={wrong}
                          title={wrong ? t('cf.manager.badRange') : undefined}
                          onFocus={() => setPicked(index)}
                          onChange={(e) => setRanges({ ...ranges, [index]: e.target.value })}
                        />
                      </td>
                      <td>
                        <input
                          type="checkbox"
                          aria-label={`${t('cf.manager.stopIfTrue')}: ${name}`}
                          checked={rule.stopIfTrue}
                          disabled={!!(rule.colorScale || rule.dataBar || rule.iconSet)}
                          onChange={(e) => setRules(m.rules.map((r, i) => (i === index ? { ...r, stopIfTrue: e.target.checked } : r)))}
                        />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          ) : (
            <p>{t('cf.manager.none')}</p>
          )}
          {bad && (
            <p className="wbe-field-error" role="alert">
              {t('cf.manager.badRange')}
            </p>
          )}
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="button" className="wbe-btn wbe-primary" onClick={apply} disabled={bad}>
              {t('dialog.ok')}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/** How a gallery's item looks: a bar of its colour, its scale's colours, or its icons' set. */
function swatch(pick: { color?: string; colors?: string[] }): CSSProperties {
  if (pick.color) return { background: `linear-gradient(to right, ${pick.color}, transparent)` };
  if (pick.colors) return { background: `linear-gradient(to right, ${pick.colors.join(', ')})` };
  return {};
}

/** Excel's Data Bars, Color Scales and Icon Sets galleries: picking one gives the selection a rule of that kind. */
export function CfGallery({ store }: { store: WorkbookStore }) {
  const t = useT();
  const at = store.cfGallery;
  const id = useId();
  const dialog = useRef<HTMLDivElement>(null);
  const close = () => store.set('cfGallery', undefined);
  useDialog(dialog, !!at, close, { modal: true });
  if (!at) return null;
  const pick = (p: Parameters<typeof galleryRule>[1]) => {
    close();
    void store.addConditionalFormat(galleryRule(at.rects, p));
  };
  const group = (title: MessageKey, items: { key: MessageKey; color?: string; colors?: string[]; set?: string }[], icon?: boolean) => (
    <>
      <h3>{t(title)}</h3>
      {items.map((p) => (
        <button key={p.key} type="button" className="wbe-btn wbe-ghost" onClick={() => pick(p as Parameters<typeof galleryRule>[1])}>
          {!icon && <span className="wbe-cf-swatch" style={swatch(p)} aria-hidden="true" />}
          {t(p.key)}
        </button>
      ))}
    </>
  );
  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <div ref={dialog} className="wbe-modal wbe-modal-wide" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1} onMouseDown={(e) => e.stopPropagation()}>
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('cf.menu.gallery')}</h2>
          <div className="wbe-cf-gallery">
            {group('cf.menu.dataBars', CF_GALLERY.dataBars)}
            {group('cf.menu.colorScales', CF_GALLERY.colorScales)}
            {group('cf.menu.iconSets', CF_GALLERY.iconSets, true)}
          </div>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

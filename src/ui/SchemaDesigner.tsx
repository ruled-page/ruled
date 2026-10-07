import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { parseA1, toA1 } from '../core/refs';
import { VT } from '../core/types';
import { bindSchema } from '../schema/binding';
import { generateSchema, inferField, inferTable } from '../schema/generate';
import { matchLabel, schemaFromJSON, schemaToJSON, textMatches } from '../schema/json';
import type { CellSchema, DeclarativeRule, FieldSchema, FieldType, SheetSchema, TableSchema, TextMatch, WorkbookSchema } from '../schema/types';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { Tabs } from './Tabs';
import { useDialog } from './useDialog';
import type { Translator } from '../i18n';
import { errorText, msg, type Localized, type MessageKey, type Text } from '../i18n/message';
import { useT } from '../i18n/react';

const TYPES: FieldType[] = ['text', 'number', 'integer', 'percent', 'date', 'boolean', 'enum'];

const TYPE_LABELS: Record<FieldType, MessageKey> = {
  text: 'designer.type.text',
  number: 'designer.type.number',
  integer: 'designer.type.integer',
  percent: 'designer.type.percent',
  date: 'designer.type.date',
  boolean: 'designer.type.boolean',
  enum: 'designer.type.enum',
};

function typeOptions(t: Translator) {
  return TYPES.map((type) => (
    <option key={type} value={type}>
      {t(TYPE_LABELS[type])}
    </option>
  ));
}

/** True when the schema carries code hooks that JSON can't hold. */
function hasHooks(schema: WorkbookSchema): boolean {
  let found = false;
  JSON.stringify(schema, (_k, v: unknown) => {
    if (typeof v === 'function') found = true;
    return v;
  });
  return found;
}

/** A plain-JSON copy that can be edited freely. */
function toDraft(schema: WorkbookSchema): WorkbookSchema {
  return JSON.parse(schemaToJSON(schema)) as WorkbookSchema;
}

function download(name: string, text: string): void {
  const url = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** "/^Q\d$/" in a label box means a pattern; anything else is literal text. */
function parseLabel(text: string): TextMatch {
  const m = /^\/(.+)\/([a-z]*)$/.exec(text.trim());
  return m ? { pattern: m[1], flags: m[2] || undefined } : text;
}

function numberOrUndefined(v: string): number | undefined {
  return v.trim() === '' || Number.isNaN(Number(v)) ? undefined : Number(v);
}

function boundOrUndefined(v: string, type: FieldType): number | string | undefined {
  if (v.trim() === '') return undefined;
  if (type === 'date') return v.trim();
  if (type === 'percent' && v.trim().endsWith('%')) return Number(v.trim().slice(0, -1)) / 100;
  return numberOrUndefined(v);
}

function boundText(b: number | string | undefined, type: FieldType): string {
  if (b === undefined) return '';
  if (type === 'percent' && typeof b === 'number') return `${+(b * 100).toPrecision(12)}%`;
  return String(b);
}

/**
 * Host-written text after it is edited in `language`. A plain string is
 * replaced. In a map of languages only that language's entry changes (under
 * the key the map already uses for it, else the base language, such as "de"),
 * and the other languages are kept.
 */
function editLocalized(value: Localized | undefined, language: string, text: string): Localized {
  if (value === undefined || typeof value === 'string') return text;
  const lng = language.toLowerCase();
  const base = lng.split('-')[0];
  const keys = Object.keys(value);
  const key = keys.find((k) => k.toLowerCase() === lng) ?? keys.find((k) => k.toLowerCase() === base) ?? base;
  return { ...value, [key]: text };
}

/** True when host-written text is empty in every language it has. */
function isBlank(value: Localized): boolean {
  return typeof value === 'string' ? value === '' : Object.values(value).every((v) => v === '');
}

interface FieldRowProps {
  field: FieldSchema;
  onChange: (f: FieldSchema) => void;
  onRemove: () => void;
}

function FieldRow({ field: f, onChange, onRemove }: FieldRowProps) {
  const t = useT();
  const set = <K extends keyof FieldSchema>(k: K, v: FieldSchema[K]) => {
    const next = { ...f, [k]: v };
    if (v === undefined || v === false || v === '') delete next[k];
    onChange(next);
  };
  const numeric = f.type === 'number' || f.type === 'integer' || f.type === 'percent';
  return (
    <tr>
      <td>
        <input value={matchLabel(f.label)} onChange={(e) => set('label', parseLabel(e.target.value))} aria-label={t('designer.field.label')} />
      </td>
      <td>
        <input className="mono" value={f.key} onChange={(e) => set('key', e.target.value.replace(/\s+/g, ''))} aria-label={t('designer.field.key')} />
      </td>
      <td>
        <select value={f.type} onChange={(e) => set('type', e.target.value as FieldType)} aria-label={t('designer.field.type')}>
          {typeOptions(t)}
        </select>
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.required} onChange={(e) => set('required', e.target.checked)} aria-label={t('designer.field.required')} />
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.unique} onChange={(e) => set('unique', e.target.checked)} aria-label={t('designer.field.unique')} />
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.readOnly} onChange={(e) => set('readOnly', e.target.checked)} aria-label={t('designer.field.readOnly')} />
      </td>
      <td>
        <select value={f.formula ?? 'allow'} onChange={(e) => set('formula', e.target.value === 'allow' ? undefined : (e.target.value as FieldSchema['formula']))} aria-label={t('designer.field.formulas')}>
          <option value="allow">{t('designer.field.formulaAllow')}</option>
          <option value="require">{t('designer.field.formulaRequire')}</option>
          <option value="forbid">{t('designer.field.formulaForbid')}</option>
        </select>
      </td>
      {f.type === 'text' ? (
        <>
          <td>
            <input className="narrow" value={f.minLength ?? ''} placeholder={t('designer.field.lengthPlaceholder')} onChange={(e) => set('minLength', numberOrUndefined(e.target.value))} aria-label={t('designer.field.minLength')} title={t('designer.field.minLengthTitle')} />
          </td>
          <td>
            <input className="narrow" value={f.maxLength ?? ''} placeholder={t('designer.field.lengthPlaceholder')} onChange={(e) => set('maxLength', numberOrUndefined(e.target.value))} aria-label={t('designer.field.maxLength')} title={t('designer.field.maxLengthTitle')} />
          </td>
        </>
      ) : (
        <>
          <td>
            <input className="narrow" value={boundText(f.min, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? t('designer.field.datePlaceholder') : ''} onChange={(e) => set('min', boundOrUndefined(e.target.value, f.type))} aria-label={t('designer.field.minimum')} />
          </td>
          <td>
            <input className="narrow" value={boundText(f.max, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? t('designer.field.datePlaceholder') : ''} onChange={(e) => set('max', boundOrUndefined(e.target.value, f.type))} aria-label={t('designer.field.maximum')} />
          </td>
        </>
      )}
      <td>
        <input className="narrow" value={f.decimals ?? ''} disabled={!numeric} onChange={(e) => set('decimals', numberOrUndefined(e.target.value))} aria-label={t('designer.field.decimals')} />
      </td>
      <td>
        {f.type === 'enum' ? (
          <input value={(f.values ?? []).join(', ')} placeholder={t('designer.field.valuesPlaceholder')} onChange={(e) => set('values', e.target.value.split(',').map((v) => v.trim()).filter(Boolean))} aria-label={t('designer.field.values')} />
        ) : (
          <input className="mono" value={f.pattern ?? ''} placeholder={t('designer.field.patternPlaceholder')} disabled={f.type === 'boolean'} onChange={(e) => set('pattern', e.target.value || undefined)} aria-label={t('designer.field.pattern')} />
        )}
      </td>
      <td>
        <button className="btn icon ghost" onClick={onRemove} aria-label={t('designer.field.remove')} title={t('designer.field.remove')}>
          <Icon name="trash" size={14} />
        </button>
      </td>
    </tr>
  );
}

type CompareOp = Extract<DeclarativeRule, { rule: 'compare' }>['op'];
const OPS: CompareOp[] = ['<', '<=', '>', '>=', '=', '<>'];

const MONOTONIC: Record<Extract<DeclarativeRule, { rule: 'monotonic' }>['direction'], MessageKey> = {
  nonIncreasing: 'designer.checks.line.nonIncreasing',
  nonDecreasing: 'designer.checks.line.nonDecreasing',
  increasing: 'designer.checks.line.increasing',
  decreasing: 'designer.checks.line.decreasing',
};

/** One line describing a check, as the designer lists it. */
function describeCheck(c: DeclarativeRule, t: Translator): string {
  const field = c.field === '*' ? t('designer.checks.everyField') : c.field;
  const value = (v: number | string | boolean) => (typeof v === 'number' ? t.number(v, { maximumFractionDigits: 15 }) : String(v));
  switch (c.rule) {
    case 'monotonic':
      return t(MONOTONIC[c.direction], { field });
    case 'firstEquals':
      return t('designer.checks.line.firstEquals', { field, value: value(c.value) });
    case 'sum': {
      // Comparison signs read the same in every language.
      const bounds = [
        c.equals !== undefined ? `= ${value(c.equals)}` : '',
        c.max !== undefined ? `≤ ${value(c.max)}` : '',
        c.min !== undefined ? `≥ ${value(c.min)}` : '',
      ];
      return t('designer.checks.line.sum', { field, bounds: bounds.filter(Boolean).join(' ') });
    }
    case 'compare':
      return t('designer.checks.line.compare', { field, op: c.op, other: c.other });
    case 'requiredIf':
      return c.equals === undefined
        ? t('designer.checks.line.requiredIfFilled', { field, when: c.when })
        : t('designer.checks.line.requiredIfEquals', { field, when: c.when, value: value(c.equals) });
    case 'uniqueTogether':
      return t('designer.checks.line.uniqueTogether', { fields: c.fields.join(' + ') });
  }
}

function ChecksEditor({ table, onChange }: { table: TableSchema; onChange: (t: TableSchema) => void }) {
  const t = useT();
  const [kind, setKind] = useState<DeclarativeRule['rule']>('monotonic');
  const [field, setField] = useState('*');
  const [other, setOther] = useState(table.fields[1]?.key ?? table.fields[0]?.key ?? '');
  const [op, setOp] = useState<CompareOp>('<=');
  const [param, setParam] = useState('nonIncreasing');
  const checks = table.checks ?? [];
  const perRecord = kind === 'compare' || kind === 'requiredIf' || kind === 'uniqueTogether';
  const target = perRecord && field === '*' ? (table.fields[0]?.key ?? '') : field;
  const add = () => {
    let c: DeclarativeRule;
    if (kind === 'monotonic') c = { rule: 'monotonic', field, direction: param as 'nonIncreasing' };
    else if (kind === 'firstEquals') c = { rule: 'firstEquals', field, value: Number.isNaN(Number(param)) ? param : Number(param) };
    else if (kind === 'sum') c = { rule: 'sum', field, max: Number(param) };
    else if (kind === 'compare') c = { rule: 'compare', field: target, op, other };
    else if (kind === 'requiredIf') c = param.trim() ? { rule: 'requiredIf', field: target, when: other, equals: Number.isNaN(Number(param)) ? param.trim() : Number(param) } : { rule: 'requiredIf', field: target, when: other };
    else c = { rule: 'uniqueTogether', field: target, fields: [target, other] };
    onChange({ ...table, checks: [...checks, c] });
  };
  const fieldOptions = table.fields.map((f) => (
    <option key={f.key} value={f.key}>
      {f.key}
    </option>
  ));
  return (
    <div className="checks">
      <h5>{t('designer.checks.heading')}</h5>
      {checks.length === 0 && <p className="muted small">{t('designer.checks.none')}</p>}
      <ul>
        {checks.map((c, i) => (
          <li key={i}>
            <span>{describeCheck(c, t)}</span>
            <button className="btn icon ghost" aria-label={t('designer.checks.remove')} onClick={() => onChange({ ...table, checks: checks.filter((_, j) => j !== i) })}>
              <Icon name="trash" size={14} />
            </button>
          </li>
        ))}
      </ul>
      <div className="add-check">
        <select
          value={kind}
          onChange={(e) => {
            const k = e.target.value as DeclarativeRule['rule'];
            setKind(k);
            setParam(k === 'monotonic' ? 'nonIncreasing' : k === 'firstEquals' ? '100' : k === 'sum' ? '0' : '');
          }}
          aria-label={t('designer.checks.kind')}
        >
          <option value="monotonic">{t('designer.checks.kindMonotonic')}</option>
          <option value="firstEquals">{t('designer.checks.kindFirstEquals')}</option>
          <option value="sum">{t('designer.checks.kindSum')}</option>
          <option value="compare">{t('designer.checks.kindCompare')}</option>
          <option value="requiredIf">{t('designer.checks.kindRequiredIf')}</option>
          <option value="uniqueTogether">{t('designer.checks.kindUniqueTogether')}</option>
        </select>
        <select value={perRecord ? target : field} onChange={(e) => setField(e.target.value)} aria-label={t('designer.checks.field')}>
          {!perRecord && <option value="*">{t('designer.checks.everyField')}</option>}
          {fieldOptions}
        </select>
        {kind === 'monotonic' && (
          <select value={param} onChange={(e) => setParam(e.target.value)} aria-label={t('designer.checks.direction')}>
            <option value="nonIncreasing">{t('designer.checks.nonIncreasing')}</option>
            <option value="nonDecreasing">{t('designer.checks.nonDecreasing')}</option>
            <option value="increasing">{t('designer.checks.increasing')}</option>
            <option value="decreasing">{t('designer.checks.decreasing')}</option>
          </select>
        )}
        {(kind === 'firstEquals' || kind === 'sum') && <input className="narrow" value={param} onChange={(e) => setParam(e.target.value)} aria-label={t('designer.checks.value')} />}
        {kind === 'compare' && (
          <select value={op} onChange={(e) => setOp(e.target.value as CompareOp)} aria-label={t('designer.checks.comparison')}>
            {OPS.map((o) => (
              <option key={o}>{o}</option>
            ))}
          </select>
        )}
        {kind === 'requiredIf' && <span className="muted small">{t('designer.checks.when')}</span>}
        {kind === 'uniqueTogether' && <span className="muted small">{t('designer.checks.with')}</span>}
        {perRecord && (
          <select value={other} onChange={(e) => setOther(e.target.value)} aria-label={kind === 'requiredIf' ? t('designer.checks.whenField') : t('designer.checks.otherField')}>
            {fieldOptions}
          </select>
        )}
        {kind === 'requiredIf' && <input className="narrow" value={param} placeholder={t('designer.checks.anyValue')} onChange={(e) => setParam(e.target.value)} aria-label={t('designer.checks.equals')} />}
        <button className="btn" onClick={add} disabled={perRecord && target === other}>
          <Icon name="plus" size={14} /> {t('designer.checks.add')}
        </button>
      </div>
    </div>
  );
}

/** The sheet's single cells (a form beside or above its tables). */
function CellsEditor({ cells, onChange }: { cells: CellSchema[]; onChange: (cells: CellSchema[]) => void }) {
  const t = useT();
  if (!cells.length) return null;
  const set = (i: number, patch: Partial<CellSchema>) => {
    const next = { ...cells[i], ...patch } as CellSchema;
    for (const k of Object.keys(patch) as (keyof CellSchema)[]) if (patch[k] === undefined || patch[k] === false || patch[k] === '') delete next[k];
    onChange(cells.map((c, j) => (j === i ? next : c)));
  };
  return (
    <div className="table-block">
      <h5>{t('designer.cells.heading')}</h5>
      <div className="fields-wrap">
        <table className="fields">
          <thead>
            <tr>
              <th>{t('designer.cells.cell')}</th>
              <th>{t('designer.cells.labelCell')}</th>
              <th>{t('designer.field.label')}</th>
              <th>{t('designer.field.key')}</th>
              <th>{t('designer.field.type')}</th>
              <th title={t('designer.field.required')}>{t('designer.field.requiredShort')}</th>
              <th>{t('designer.field.min')}</th>
              <th>{t('designer.field.max')}</th>
              <th>{t('designer.field.valuesOrPattern')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {cells.map((c, i) => {
              const numeric = c.type === 'number' || c.type === 'integer' || c.type === 'percent';
              return (
                <tr key={i}>
                  <td>
                    <input className="mono narrow" value={c.ref} onChange={(e) => set(i, { ref: e.target.value.toUpperCase() })} aria-label={t('designer.cells.cell')} />
                  </td>
                  <td>
                    <input className="mono narrow" value={c.label?.ref ?? ''} onChange={(e) => set(i, { label: e.target.value ? { ref: e.target.value.toUpperCase(), text: c.label?.text ?? '' } : undefined })} aria-label={t('designer.cells.labelCell')} />
                  </td>
                  <td>
                    <input value={c.label ? matchLabel(c.label.text) : ''} disabled={!c.label} onChange={(e) => set(i, { label: { ref: c.label!.ref, text: parseLabel(e.target.value) } })} aria-label={t('designer.cells.labelText')} />
                  </td>
                  <td>
                    <input className="mono" value={c.key} onChange={(e) => set(i, { key: e.target.value.replace(/\s+/g, '') })} aria-label={t('designer.field.key')} />
                  </td>
                  <td>
                    <select value={c.type} onChange={(e) => set(i, { type: e.target.value as FieldType })} aria-label={t('designer.field.type')}>
                      {typeOptions(t)}
                    </select>
                  </td>
                  <td className="center">
                    <input type="checkbox" checked={!!c.required} onChange={(e) => set(i, { required: e.target.checked })} aria-label={t('designer.field.required')} />
                  </td>
                  <td>
                    <input className="narrow" value={boundText(c.min, c.type)} disabled={!numeric && c.type !== 'date'} onChange={(e) => set(i, { min: boundOrUndefined(e.target.value, c.type) })} aria-label={t('designer.field.minimum')} />
                  </td>
                  <td>
                    <input className="narrow" value={boundText(c.max, c.type)} disabled={!numeric && c.type !== 'date'} onChange={(e) => set(i, { max: boundOrUndefined(e.target.value, c.type) })} aria-label={t('designer.field.maximum')} />
                  </td>
                  <td>
                    {c.type === 'enum' ? (
                      <input value={(c.values ?? []).join(', ')} placeholder={t('designer.field.valuesPlaceholder')} onChange={(e) => set(i, { values: e.target.value.split(',').map((v) => v.trim()).filter(Boolean) })} aria-label={t('designer.field.values')} />
                    ) : (
                      <input className="mono" value={c.pattern ?? ''} placeholder={t('designer.field.patternPlaceholder')} disabled={c.type === 'boolean'} onChange={(e) => set(i, { pattern: e.target.value || undefined })} aria-label={t('designer.field.pattern')} />
                    )}
                  </td>
                  <td>
                    <button className="btn icon ghost" onClick={() => onChange(cells.filter((_, j) => j !== i))} aria-label={t('designer.cells.remove')} title={t('designer.cells.remove')}>
                      <Icon name="trash" size={14} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/**
 * While the designer waits for a selection on the sheet it shrinks to this
 * bar, so the grid can be used; adding reads the table or cell from what is selected.
 */
function PickBar({ store, draft, kind }: { store: WorkbookStore; draft: WorkbookSchema; kind: 'table' | 'cell' }) {
  const t = useT();
  const view = store.view;
  const sel = store.selection;
  const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
  // The sheet is where the picking happens.
  useEffect(() => store.focusGrid(), [store]);
  const add = () => {
    if (!view || view.info.kind !== 'worksheet') return;
    const d = structuredClone(draft);
    let sheet = d.sheets.find((s) => textMatches(s.name, view.info.name));
    if (!sheet) {
      sheet = { name: view.info.name };
      d.sheets.push(sheet);
    }
    if (kind === 'table') {
      const keys = new Set(d.sheets.flatMap((s) => (s.tables ?? []).map((tb) => tb.key)));
      const base = view.info.name.toLowerCase().replace(/[^a-z0-9]+/g, '') || 'table';
      let key = base;
      for (let n = 2; keys.has(key); n++) key = `${base}${n}`;
      const table = inferTable(view, store.styles, key, { row: rect.r1, col: rect.c1 }, 'rows', false, store.sheets);
      if (!table.fields.length) {
        store.toast('warn', msg('designer.pick.noHeadings', { ref: toA1(rect.r1, rect.c1) }));
        return;
      }
      sheet.tables = [...(sheet.tables ?? []), table];
    } else {
      const left = sel.col > 0 ? view.get(sel.row, sel.col - 1) : null;
      const labelText = left?.t === VT.String ? (left.s ?? '').trim() : '';
      const used = new Set((sheet.cells ?? []).map((c) => c.key));
      const { label: _label, unique: _unique, ...field } = inferField(labelText.replace(/:$/, '') || toA1(sel.row, sel.col), [view.get(sel.row, sel.col)], store.styles, used);
      const cell: CellSchema = { ...field, ref: toA1(sel.row, sel.col) };
      if (labelText) cell.label = { ref: toA1(sel.row, sel.col - 1), text: labelText };
      sheet.cells = [...(sheet.cells ?? []).filter((c) => c.ref !== cell.ref), cell];
    }
    store.set('schemaDraft', d);
    store.set('designerPick', undefined);
  };
  return (
    <div
      className="pick-bar"
      role="dialog"
      aria-label={t('designer.pick.label')}
      onKeyDown={(e) => {
        if (e.key === 'Escape') store.set('designerPick', undefined);
      }}
    >
      <Icon name="schema" />
      <span>
        {kind === 'table' ? t('designer.pick.tableHint') : t('designer.pick.cellHint')}{' '}
        <span className="sr-only">{t('designer.pick.f6')}</span>
      </span>
      <button className="btn primary" onClick={add}>
        {kind === 'table' ? t('designer.pick.addTable', { ref: toA1(rect.r1, rect.c1) }) : t('designer.pick.addCell', { ref: toA1(sel.row, sel.col) })}
      </button>
      <button className="btn" onClick={() => store.set('designerPick', undefined)}>
        {t('designer.pick.back')}
      </button>
    </div>
  );
}

export function SchemaDesigner({ store }: { store: WorkbookStore }) {
  const t = useT();
  // The draft lives in the store, so it survives closing the designer or picking cells on the sheet.
  const draft = store.schemaDraft ?? null;
  const setDraft = (d: WorkbookSchema | null) => store.set('schemaDraft', d ?? undefined);
  useEffect(() => {
    if (!store.schemaDraft && store.schema) store.set('schemaDraft', toDraft(store.schema.schema));
  }, [store]);
  const [originalHadHooks, setOriginalHadHooks] = useState(() => (store.schema ? hasHooks(store.schema.schema) : false));
  const [mode, setMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState('');
  const [jsonError, setJsonError] = useState<Text>('');
  const [testResult, setTestResult] = useState<{ ok: boolean; problems: Text[] } | null>(null);
  const importInput = useRef<HTMLInputElement>(null);
  const dialog = useRef<HTMLDivElement>(null);
  const titleId = useId();
  const ready = store.status === 'ready';

  const edit = (fn: (d: WorkbookSchema) => void) => {
    if (!draft) return;
    const next = structuredClone(draft);
    fn(next);
    setDraft(next);
    setTestResult(null);
  };

  const viewFor = (sheet: SheetSchema) => store.sheets.find((v) => v.info.kind === 'worksheet' && textMatches(sheet.name, v.info.name));

  const reread = (si: number, ti: number, changes: Partial<Pick<TableSchema, 'orientation' | 'at'>> & { fixed?: boolean }) => {
    if (!draft) return;
    const sheet = draft.sheets[si];
    const table = sheet.tables![ti];
    const view = viewFor(sheet);
    const at = parseA1(changes.at ?? table.at);
    if (!view || !at) {
      edit((d) => Object.assign(d.sheets[si].tables![ti], changes));
      return;
    }
    const fixed = changes.fixed ?? !!table.records?.labels;
    const next = inferTable(view, store.styles, table.key, at, changes.orientation ?? table.orientation ?? 'rows', fixed, store.sheets);
    edit((d) => {
      d.sheets[si].tables![ti] = { ...next, checks: table.checks };
    });
  };

  const startFrom = (schema: WorkbookSchema) => {
    setDraft(toDraft(schema));
    setOriginalHadHooks(hasHooks(schema));
    setMode('form');
    setTestResult(null);
    setJsonError('');
  };

  const generate = () => {
    const schema = generateSchema(store.sheets, store.styles, store.fileName || 'workbook');
    if (!schema.sheets.length) {
      store.toast('warn', msg('designer.toast.noTables'));
      return;
    }
    startFrom(schema);
  };

  const onImport = async (file: File | undefined) => {
    if (!file) return;
    try {
      const schema = schemaFromJSON(await file.text());
      store.registerSchema(schema);
      startFrom(schema);
      const applied = store.schema?.schema.id === schema.id;
      const name = t.localized(schema.name);
      store.toast(applied ? 'info' : 'warn', msg(applied ? 'designer.toast.importedApplied' : 'designer.toast.importedNoMatch', { name }));
    } catch (e) {
      store.toast('error', errorText(e));
    }
  };

  const currentDraft = (): WorkbookSchema | null => {
    if (mode === 'form') return draft;
    try {
      const parsed = schemaFromJSON(jsonText);
      setJsonError('');
      return parsed;
    } catch (e) {
      setJsonError(errorText(e));
      return null;
    }
  };

  const test = () => {
    const d = currentDraft();
    if (!d) return;
    const r = bindSchema(d, store.sheets);
    setTestResult({ ok: !!r.binding, problems: r.problems });
  };

  const apply = () => {
    const d = currentDraft();
    if (!d) return;
    try {
      schemaFromJSON(schemaToJSON(d));
    } catch (e) {
      store.toast('error', errorText(e));
      return;
    }
    store.registerSchema(d);
    setDraft(d);
    setOriginalHadHooks(false);
    const applied = store.schema?.schema.id === d.id;
    if (applied) {
      store.set('designerOpen', false);
      store.set('panel', 'issues');
      store.toast('info', msg('designer.toast.applied', { name: t.localized(d.name) }));
    } else {
      test();
    }
  };

  const json = useMemo(() => (draft ? schemaToJSON(draft) : ''), [draft]);
  const activeId = store.schema?.schema.id;

  const picking = !!(store.designerPick && draft);
  useDialog(dialog, !picking, () => store.set('designerOpen', false), { modal: true });
  if (picking) return <PickBar store={store} draft={draft!} kind={store.designerPick!} />;

  return (
    <div className="modal-backdrop" onMouseDown={() => store.set('designerOpen', false)}>
      <div
        ref={dialog}
        className="designer"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${titleId}-title`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <header className="designer-head">
          <h2 id={`${titleId}-title`}>
            <Icon name="schema" /> {t('designer.title')}
          </h2>
          <button className="btn icon ghost" onClick={() => store.set('designerOpen', false)} aria-label={t('designer.close')}>
            <Icon name="close" />
          </button>
        </header>
        <div className="designer-body">
          <nav className="designer-list">
            <h3>{t('designer.list.heading')}</h3>
            <p className="muted small">{t('designer.list.intro')}</p>
            <ul>
              {store.schemas.map((s) => (
                <li key={s.id} className={draft?.id === s.id ? 'selected' : ''}>
                  <button className="schema-item" onClick={() => startFrom(s)}>
                    <strong>{t.localized(s.name)}</strong>
                    <span className="mono small">{s.id}</span>
                    <span className="tags">
                      {s.id === activeId && <span className="tag tag-ok">{t('designer.list.applied')}</span>}
                      {hasHooks(s) && <span className="tag">{t('designer.list.codeHooks')}</span>}
                    </span>
                  </button>
                  <div className="item-actions">
                    <button className="btn icon ghost" title={t('designer.exportJson')} aria-label={t('designer.list.export', { name: t.localized(s.name) })} onClick={() => download(`${s.id}.schema.json`, schemaToJSON(s))}>
                      <Icon name="download" size={14} />
                    </button>
                    <button className="btn icon ghost" title={t('designer.list.removeTitle')} aria-label={t('designer.list.remove', { name: t.localized(s.name) })} onClick={() => store.removeSchema(s.id)}>
                      <Icon name="trash" size={14} />
                    </button>
                  </div>
                </li>
              ))}
            </ul>
            <div className="list-actions">
              <button className="btn primary" disabled={!ready} onClick={generate} title={ready ? '' : t('designer.list.generateNeedsWorkbook')}>
                <Icon name="sparkle" /> {t('designer.list.generate')}
              </button>
              <button className="btn" onClick={() => importInput.current?.click()}>
                <Icon name="upload" /> {t('designer.list.import')}
              </button>
              <input ref={importInput} type="file" accept=".json,application/json" hidden onChange={(e) => {
                void onImport(e.target.files?.[0]);
                e.target.value = '';
              }} />
            </div>
          </nav>

          <section className="designer-editor">
            {!draft && (
              <div className="center-state static">
                <Icon name="schema" size={28} />
                <p>{t('designer.empty')}</p>
              </div>
            )}
            {draft && (
              <>
                <Tabs
                  label={t('designer.editAs')}
                  tabs={[
                    { id: 'form', label: t('designer.tab.form') },
                    { id: 'json', label: t('designer.tab.json') },
                  ]}
                  value={mode}
                  onChange={(next) => {
                    if (next === 'json' && mode !== 'json') {
                      setJsonText(json);
                      setJsonError('');
                    }
                    setMode(next);
                  }}
                  head={(tabs) => (
                    <div className="editor-tabs">
                      {tabs}
                      {originalHadHooks && (
                        <p className="hint warn">{t('designer.hooksWarning')}</p>
                      )}
                    </div>
                  )}
                >
                  {mode === 'json' ? (
                    <div className="json-editor">
                      <textarea className="mono" value={jsonText} spellCheck={false} onChange={(e) => setJsonText(e.target.value)} aria-label={t('designer.json.label')} />
                      {jsonError && <p className="hint error">{t.text(jsonError)}</p>}
                    </div>
                  ) : (
                    <div className="form">
                      <div className="grid-form">
                        <label>
                          {t('designer.schema.name')}
                          <input value={t.localized(draft.name)} onChange={(e) => edit((d) => void (d.name = editLocalized(d.name, t.language, e.target.value)))} />
                        </label>
                        <label>
                          {t('designer.schema.id')}
                          <input className="mono" value={draft.id} onChange={(e) => edit((d) => void (d.id = e.target.value.replace(/\s+/g, '-')))} />
                        </label>
                        <label>
                          {t('designer.schema.version')}
                          <input className="mono" value={draft.version ?? ''} placeholder={t('designer.schema.versionPlaceholder')} onChange={(e) => edit((d) => void (d.version = e.target.value.trim() || undefined))} title={t('designer.schema.versionTitle')} />
                        </label>
                        <label>
                          {t('designer.schema.hookTimeout')}
                          <input className="narrow" value={draft.hookTimeoutMs === undefined ? '' : draft.hookTimeoutMs / 1000} placeholder="10" onChange={(e) => edit((d) => {
                            const n = numberOrUndefined(e.target.value);
                            d.hookTimeoutMs = n === undefined ? undefined : Math.round(n * 1000);
                          })} />
                        </label>
                        <label className="wide">
                          {t('designer.schema.description')}
                          <input
                            value={t.localized(draft.description)}
                            onChange={(e) => edit((d) => {
                              const next = editLocalized(d.description, t.language, e.target.value);
                              d.description = isBlank(next) ? undefined : next;
                            })}
                          />
                        </label>
                        <label>
                          {t('designer.schema.otherSheets')}
                          <select value={draft.otherSheets ?? 'editable'} onChange={(e) => edit((d) => void (d.otherSheets = e.target.value as 'editable'))}>
                            <option value="editable">{t('designer.schema.editable')}</option>
                            <option value="readOnly">{t('designer.schema.readOnly')}</option>
                          </select>
                        </label>
                        <label>
                          {t('designer.schema.saveWithErrors')}
                          <select value={draft.saveWithErrors ?? 'block'} onChange={(e) => edit((d) => void (d.saveWithErrors = e.target.value as 'block'))}>
                            <option value="block">{t('designer.schema.saveBlocked')}</option>
                            <option value="confirm">{t('designer.schema.saveConfirm')}</option>
                          </select>
                        </label>
                        <label>
                          {t('designer.schema.numbers')}
                          <select
                            value={draft.numbers?.display === 'full' ? 'exact' : 'formatted'}
                            onChange={(e) => edit((d) => void (d.numbers = e.target.value === 'exact' ? { display: 'full', copy: 'raw' } : undefined))}
                          >
                            <option value="formatted">{t('designer.schema.numbersFormatted')}</option>
                            <option value="exact">{t('designer.schema.numbersExact')}</option>
                          </select>
                        </label>
                      </div>

                      {draft.sheets.map((sheet, si) => (
                        <div key={si} className="sheet-block">
                          <div className="sheet-head">
                            <label>
                              {t('designer.sheet.name')}
                              <input value={matchLabel(sheet.name)} onChange={(e) => edit((d) => void (d.sheets[si].name = parseLabel(e.target.value)))} />
                            </label>
                            <label>
                              {t('designer.sheet.outside')}
                              <select value={sheet.outside ?? 'readOnly'} onChange={(e) => edit((d) => void (d.sheets[si].outside = e.target.value as 'readOnly'))}>
                                <option value="readOnly">{t('designer.schema.readOnly')}</option>
                                <option value="editable">{t('designer.schema.editable')}</option>
                              </select>
                            </label>
                            <label className="inline">
                              <input type="checkbox" checked={!!sheet.optional} onChange={(e) => edit((d) => void (d.sheets[si].optional = e.target.checked || undefined))} /> {t('designer.sheet.optional')}
                            </label>
                            {!viewFor(sheet) && ready && <span className="tag tag-warn">{t('designer.sheet.missing')}</span>}
                            <button className="btn icon ghost" aria-label={t('designer.sheet.remove')} title={t('designer.sheet.remove')} onClick={() => edit((d) => void d.sheets.splice(si, 1))}>
                              <Icon name="trash" size={14} />
                            </button>
                          </div>
                          <CellsEditor
                            cells={sheet.cells ?? []}
                            onChange={(cells) => edit((d) => void (d.sheets[si].cells = cells.length ? cells : undefined))}
                          />
                          {(sheet.tables ?? []).map((table, ti) => (
                            <div key={ti} className="table-block">
                              <div className="table-head">
                                <label>
                                  {t('designer.table.key')}
                                  <input className="mono" value={table.key} onChange={(e) => edit((d) => void (d.sheets[si].tables![ti].key = e.target.value))} />
                                </label>
                                <label>
                                  {t('designer.table.at')}
                                  <input className="mono narrow" value={table.at} onChange={(e) => edit((d) => void (d.sheets[si].tables![ti].at = e.target.value.toUpperCase()))} />
                                </label>
                                <label>
                                  {t('designer.table.orientation')}
                                  <select value={table.orientation ?? 'rows'} onChange={(e) => reread(si, ti, { orientation: e.target.value as 'rows' })}>
                                    <option value="rows">{t('designer.table.rows')}</option>
                                    <option value="columns">{t('designer.table.columns')}</option>
                                  </select>
                                </label>
                                <label className="inline" title={t('designer.table.fixedTitle')}>
                                  <input type="checkbox" checked={!!table.records?.labels} onChange={(e) => reread(si, ti, { fixed: e.target.checked })} />{' '}
                                  {table.orientation === 'columns' ? t('designer.table.fixedColumns') : t('designer.table.fixedRows')}
                                </label>
                                {!table.records?.labels && (
                                  <>
                                    <label>
                                      {t('designer.table.minRecords')}
                                      <input className="narrow" value={table.records?.min ?? ''} onChange={(e) => edit((d) => {
                                        const tb = d.sheets[si].tables![ti];
                                        tb.records = { ...tb.records, min: numberOrUndefined(e.target.value) };
                                      })} />
                                    </label>
                                    <label>
                                      {t('designer.table.maxRecords')}
                                      <input className="narrow" value={table.records?.max ?? ''} onChange={(e) => edit((d) => {
                                        const tb = d.sheets[si].tables![ti];
                                        tb.records = { ...tb.records, max: numberOrUndefined(e.target.value) };
                                      })} />
                                    </label>
                                  </>
                                )}
                                <button className="btn" disabled={!viewFor(sheet)} onClick={() => reread(si, ti, {})} title={t('designer.table.rereadTitle')}>
                                  {t('designer.table.reread')}
                                </button>
                              </div>
                              {table.records?.labels && (
                                <p className="muted small">
                                  {t(table.orientation === 'columns' ? 'designer.table.fixedColumnLabels' : 'designer.table.fixedRowLabels', {
                                    labels: table.records.labels.map((l) => matchLabel(l)).join(', '),
                                  })}
                                </p>
                              )}
                              <div className="fields-wrap">
                                <table className="fields">
                                  <thead>
                                    <tr>
                                      <th>{t('designer.field.label')}</th>
                                      <th>{t('designer.field.key')}</th>
                                      <th>{t('designer.field.type')}</th>
                                      <th title={t('designer.field.required')}>{t('designer.field.requiredShort')}</th>
                                      <th title={t('designer.field.unique')}>{t('designer.field.uniqueShort')}</th>
                                      <th title={t('designer.field.readOnly')}>{t('designer.field.readOnlyShort')}</th>
                                      <th>{t('designer.field.formulas')}</th>
                                      <th>{t('designer.field.min')}</th>
                                      <th>{t('designer.field.max')}</th>
                                      <th title={t('designer.field.decimals')}>{t('designer.field.decimalsShort')}</th>
                                      <th>{t('designer.field.valuesOrPattern')}</th>
                                      <th />
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {table.fields.map((f, fi) => (
                                      <FieldRow
                                        key={fi}
                                        field={f}
                                        onChange={(nf) => edit((d) => void (d.sheets[si].tables![ti].fields[fi] = nf))}
                                        onRemove={() => edit((d) => void d.sheets[si].tables![ti].fields.splice(fi, 1))}
                                      />
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                              <ChecksEditor table={table} onChange={(next) => edit((d) => void (d.sheets[si].tables![ti] = next))} />
                              <button className="btn ghost small" onClick={() => edit((d) => void d.sheets[si].tables!.splice(ti, 1))}>
                                <Icon name="trash" size={14} /> {t('designer.table.remove')}
                              </button>
                            </div>
                          ))}
                        </div>
                      ))}
                      {ready && (
                        <div className="pick-actions">
                          <span className="muted small">
                            {store.view ? t('designer.pick.prompt', { sheet: store.view.info.name }) : t('designer.pick.promptNoSheet')}
                          </span>
                          <button className="btn" onClick={() => store.set('designerPick', 'table')}>
                            <Icon name="plus" size={14} /> {t('designer.pick.table')}
                          </button>
                          <button className="btn" onClick={() => store.set('designerPick', 'cell')}>
                            <Icon name="plus" size={14} /> {t('designer.pick.cell')}
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </Tabs>

                {testResult && (
                  <div className={`test-result ${testResult.ok ? 'ok' : 'bad'}`}>
                    <Icon name={testResult.ok ? 'check' : 'warn'} />
                    <div>
                      <strong>{testResult.ok ? t('designer.test.matches') : t('designer.test.noMatch')}</strong>
                      {testResult.problems.length > 0 && (
                        <ul>
                          {testResult.problems.map((p, i) => (
                            <li key={i}>{t.text(p)}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>
                )}
              </>
            )}
          </section>
        </div>
        {draft && (
          <footer className="designer-foot">
            <button className="btn" onClick={() => {
              const d = currentDraft();
              if (d) download(`${d.id}.schema.json`, schemaToJSON(d));
            }}>
              <Icon name="download" /> {t('designer.exportJson')}
            </button>
            <span className="spacer" />
            <button className="btn" disabled={!ready} onClick={test}>
              {t('designer.test.run')}
            </button>
            <button className="btn primary" onClick={apply}>
              <Icon name="check" /> {t('designer.apply')}
            </button>
          </footer>
        )}
      </div>
    </div>
  );
}

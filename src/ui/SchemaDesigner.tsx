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

const TYPES: FieldType[] = ['text', 'number', 'integer', 'percent', 'date', 'boolean', 'enum'];

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

interface FieldRowProps {
  field: FieldSchema;
  onChange: (f: FieldSchema) => void;
  onRemove: () => void;
}

function FieldRow({ field: f, onChange, onRemove }: FieldRowProps) {
  const set = <K extends keyof FieldSchema>(k: K, v: FieldSchema[K]) => {
    const next = { ...f, [k]: v };
    if (v === undefined || v === false || v === '') delete next[k];
    onChange(next);
  };
  const numeric = f.type === 'number' || f.type === 'integer' || f.type === 'percent';
  return (
    <tr>
      <td>
        <input value={matchLabel(f.label)} onChange={(e) => set('label', parseLabel(e.target.value))} aria-label="Label" />
      </td>
      <td>
        <input className="mono" value={f.key} onChange={(e) => set('key', e.target.value.replace(/\s+/g, ''))} aria-label="Key" />
      </td>
      <td>
        <select value={f.type} onChange={(e) => set('type', e.target.value as FieldType)} aria-label="Type">
          {TYPES.map((t) => (
            <option key={t}>{t}</option>
          ))}
        </select>
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.required} onChange={(e) => set('required', e.target.checked)} aria-label="Required" />
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.unique} onChange={(e) => set('unique', e.target.checked)} aria-label="Unique" />
      </td>
      <td className="center">
        <input type="checkbox" checked={!!f.readOnly} onChange={(e) => set('readOnly', e.target.checked)} aria-label="Read-only" />
      </td>
      <td>
        <select value={f.formula ?? 'allow'} onChange={(e) => set('formula', e.target.value === 'allow' ? undefined : (e.target.value as FieldSchema['formula']))} aria-label="Formulas">
          <option value="allow">allow</option>
          <option value="require">require</option>
          <option value="forbid">forbid</option>
        </select>
      </td>
      {f.type === 'text' ? (
        <>
          <td>
            <input className="narrow" value={f.minLength ?? ''} placeholder="length" onChange={(e) => set('minLength', numberOrUndefined(e.target.value))} aria-label="Shortest length" title="Fewest characters" />
          </td>
          <td>
            <input className="narrow" value={f.maxLength ?? ''} placeholder="length" onChange={(e) => set('maxLength', numberOrUndefined(e.target.value))} aria-label="Longest length" title="Most characters" />
          </td>
        </>
      ) : (
        <>
          <td>
            <input className="narrow" value={boundText(f.min, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? 'YYYY-MM-DD' : ''} onChange={(e) => set('min', boundOrUndefined(e.target.value, f.type))} aria-label="Minimum" />
          </td>
          <td>
            <input className="narrow" value={boundText(f.max, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? 'YYYY-MM-DD' : ''} onChange={(e) => set('max', boundOrUndefined(e.target.value, f.type))} aria-label="Maximum" />
          </td>
        </>
      )}
      <td>
        <input className="narrow" value={f.decimals ?? ''} disabled={!numeric} onChange={(e) => set('decimals', numberOrUndefined(e.target.value))} aria-label="Decimal places" />
      </td>
      <td>
        {f.type === 'enum' ? (
          <input value={(f.values ?? []).join(', ')} placeholder="a, b, c" onChange={(e) => set('values', e.target.value.split(',').map((v) => v.trim()).filter(Boolean))} aria-label="Allowed values" />
        ) : (
          <input className="mono" value={f.pattern ?? ''} placeholder="regular expression" disabled={f.type === 'boolean'} onChange={(e) => set('pattern', e.target.value || undefined)} aria-label="Pattern" />
        )}
      </td>
      <td>
        <button className="btn icon ghost" onClick={onRemove} aria-label="Remove field" title="Remove field">
          <Icon name="trash" size={14} />
        </button>
      </td>
    </tr>
  );
}

type CompareOp = Extract<DeclarativeRule, { rule: 'compare' }>['op'];
const OPS: CompareOp[] = ['<', '<=', '>', '>=', '=', '<>'];

/** One line describing a check, as the designer lists it. */
function describeCheck(c: DeclarativeRule): string {
  const f = c.field === '*' ? 'every field' : c.field;
  switch (c.rule) {
    case 'monotonic':
      return `${f}: ${c.direction} down the table`;
    case 'firstEquals':
      return `${f}: first value is ${c.value}`;
    case 'sum':
      return `${f}: total ${c.equals !== undefined ? `= ${c.equals}` : ''}${c.max !== undefined ? ` ≤ ${c.max}` : ''}${c.min !== undefined ? ` ≥ ${c.min}` : ''}`;
    case 'compare':
      return `${f} ${c.op} ${c.other} in every record`;
    case 'requiredIf':
      return `${f}: required when ${c.when} ${c.equals === undefined ? 'is filled in' : `is ${c.equals}`}`;
    case 'uniqueTogether':
      return `${c.fields.join(' + ')}: unique together`;
  }
}

function ChecksEditor({ table, onChange }: { table: TableSchema; onChange: (t: TableSchema) => void }) {
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
      <h5>Table checks</h5>
      {checks.length === 0 && <p className="muted small">None. Field rules above still apply.</p>}
      <ul>
        {checks.map((c, i) => (
          <li key={i}>
            <span>{describeCheck(c)}</span>
            <button className="btn icon ghost" aria-label="Remove check" onClick={() => onChange({ ...table, checks: checks.filter((_, j) => j !== i) })}>
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
          aria-label="Check type"
        >
          <option value="monotonic">Values only go one way</option>
          <option value="firstEquals">First value equals</option>
          <option value="sum">Total at most</option>
          <option value="compare">Compare two fields</option>
          <option value="requiredIf">Required when another is set</option>
          <option value="uniqueTogether">Unique together</option>
        </select>
        <select value={perRecord ? target : field} onChange={(e) => setField(e.target.value)} aria-label="Field">
          {!perRecord && <option value="*">every field</option>}
          {fieldOptions}
        </select>
        {kind === 'monotonic' && (
          <select value={param} onChange={(e) => setParam(e.target.value)} aria-label="Direction">
            <option value="nonIncreasing">never go up</option>
            <option value="nonDecreasing">never go down</option>
            <option value="increasing">always go up</option>
            <option value="decreasing">always go down</option>
          </select>
        )}
        {(kind === 'firstEquals' || kind === 'sum') && <input className="narrow" value={param} onChange={(e) => setParam(e.target.value)} aria-label="Value" />}
        {kind === 'compare' && (
          <select value={op} onChange={(e) => setOp(e.target.value as CompareOp)} aria-label="Comparison">
            {OPS.map((o) => (
              <option key={o}>{o}</option>
            ))}
          </select>
        )}
        {kind === 'requiredIf' && <span className="muted small">when</span>}
        {kind === 'uniqueTogether' && <span className="muted small">with</span>}
        {perRecord && (
          <select value={other} onChange={(e) => setOther(e.target.value)} aria-label={kind === 'requiredIf' ? 'When field' : 'Other field'}>
            {fieldOptions}
          </select>
        )}
        {kind === 'requiredIf' && <input className="narrow" value={param} placeholder="has any value" onChange={(e) => setParam(e.target.value)} aria-label="Equals" />}
        <button className="btn" onClick={add} disabled={perRecord && target === other}>
          <Icon name="plus" size={14} /> Add
        </button>
      </div>
    </div>
  );
}

/** The sheet's single cells (a form beside or above its tables). */
function CellsEditor({ cells, onChange }: { cells: CellSchema[]; onChange: (cells: CellSchema[]) => void }) {
  if (!cells.length) return null;
  const set = (i: number, patch: Partial<CellSchema>) => {
    const next = { ...cells[i], ...patch } as CellSchema;
    for (const k of Object.keys(patch) as (keyof CellSchema)[]) if (patch[k] === undefined || patch[k] === false || patch[k] === '') delete next[k];
    onChange(cells.map((c, j) => (j === i ? next : c)));
  };
  return (
    <div className="table-block">
      <h5>Single cells</h5>
      <div className="fields-wrap">
        <table className="fields">
          <thead>
            <tr>
              <th>Cell</th>
              <th>Label cell</th>
              <th>Label</th>
              <th>Key</th>
              <th>Type</th>
              <th title="Required">Req</th>
              <th>Min</th>
              <th>Max</th>
              <th>Values / pattern</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {cells.map((c, i) => {
              const numeric = c.type === 'number' || c.type === 'integer' || c.type === 'percent';
              return (
                <tr key={i}>
                  <td>
                    <input className="mono narrow" value={c.ref} onChange={(e) => set(i, { ref: e.target.value.toUpperCase() })} aria-label="Cell" />
                  </td>
                  <td>
                    <input className="mono narrow" value={c.label?.ref ?? ''} onChange={(e) => set(i, { label: e.target.value ? { ref: e.target.value.toUpperCase(), text: c.label?.text ?? '' } : undefined })} aria-label="Label cell" />
                  </td>
                  <td>
                    <input value={c.label ? matchLabel(c.label.text) : ''} disabled={!c.label} onChange={(e) => set(i, { label: { ref: c.label!.ref, text: parseLabel(e.target.value) } })} aria-label="Label text" />
                  </td>
                  <td>
                    <input className="mono" value={c.key} onChange={(e) => set(i, { key: e.target.value.replace(/\s+/g, '') })} aria-label="Key" />
                  </td>
                  <td>
                    <select value={c.type} onChange={(e) => set(i, { type: e.target.value as FieldType })} aria-label="Type">
                      {TYPES.map((t) => (
                        <option key={t}>{t}</option>
                      ))}
                    </select>
                  </td>
                  <td className="center">
                    <input type="checkbox" checked={!!c.required} onChange={(e) => set(i, { required: e.target.checked })} aria-label="Required" />
                  </td>
                  <td>
                    <input className="narrow" value={boundText(c.min, c.type)} disabled={!numeric && c.type !== 'date'} onChange={(e) => set(i, { min: boundOrUndefined(e.target.value, c.type) })} aria-label="Minimum" />
                  </td>
                  <td>
                    <input className="narrow" value={boundText(c.max, c.type)} disabled={!numeric && c.type !== 'date'} onChange={(e) => set(i, { max: boundOrUndefined(e.target.value, c.type) })} aria-label="Maximum" />
                  </td>
                  <td>
                    {c.type === 'enum' ? (
                      <input value={(c.values ?? []).join(', ')} placeholder="a, b, c" onChange={(e) => set(i, { values: e.target.value.split(',').map((v) => v.trim()).filter(Boolean) })} aria-label="Allowed values" />
                    ) : (
                      <input className="mono" value={c.pattern ?? ''} placeholder="regular expression" disabled={c.type === 'boolean'} onChange={(e) => set(i, { pattern: e.target.value || undefined })} aria-label="Pattern" />
                    )}
                  </td>
                  <td>
                    <button className="btn icon ghost" onClick={() => onChange(cells.filter((_, j) => j !== i))} aria-label="Remove cell" title="Remove cell">
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
      const keys = new Set(d.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.key)));
      const base = view.info.name.toLowerCase().replace(/[^a-z0-9]+/g, '') || 'table';
      let key = base;
      for (let n = 2; keys.has(key); n++) key = `${base}${n}`;
      const table = inferTable(view, store.styles, key, { row: rect.r1, col: rect.c1 }, 'rows', false, store.sheets);
      if (!table.fields.length) {
        store.toast('warn', `There are no headings at ${toA1(rect.r1, rect.c1)}. Select from the table's heading row.`);
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
      aria-label="Pick cells for the schema"
      onKeyDown={(e) => {
        if (e.key === 'Escape') store.set('designerPick', undefined);
      }}
    >
      <Icon name="schema" />
      <span>
        {kind === 'table'
          ? 'Select the table’s heading row (or its top-left cell) on the sheet, then add it.'
          : 'Select the cell to check; a label just to its left is picked up too.'}{' '}
        <span className="sr-only">F6 moves between the sheet and this bar.</span>
      </span>
      <button className="btn primary" onClick={add}>
        Add {kind === 'table' ? `table at ${toA1(rect.r1, rect.c1)}` : `cell ${toA1(sel.row, sel.col)}`}
      </button>
      <button className="btn" onClick={() => store.set('designerPick', undefined)}>
        Back
      </button>
    </div>
  );
}

export function SchemaDesigner({ store }: { store: WorkbookStore }) {
  // The draft lives in the store, so it survives closing the designer or picking cells on the sheet.
  const draft = store.schemaDraft ?? null;
  const setDraft = (d: WorkbookSchema | null) => store.set('schemaDraft', d ?? undefined);
  useEffect(() => {
    if (!store.schemaDraft && store.schema) store.set('schemaDraft', toDraft(store.schema.schema));
  }, [store]);
  const [originalHadHooks, setOriginalHadHooks] = useState(() => (store.schema ? hasHooks(store.schema.schema) : false));
  const [mode, setMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState('');
  const [jsonError, setJsonError] = useState('');
  const [testResult, setTestResult] = useState<{ ok: boolean; problems: string[] } | null>(null);
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
      store.toast('warn', 'No table with a header row was found to build a schema from.');
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
      store.toast(applied ? 'info' : 'warn', applied ? `Imported “${schema.name}” and applied it to this workbook.` : `Imported “${schema.name}”. It doesn't match the open workbook.`);
    } catch (e) {
      store.toast('error', (e as Error).message);
    }
  };

  const currentDraft = (): WorkbookSchema | null => {
    if (mode === 'form') return draft;
    try {
      const parsed = schemaFromJSON(jsonText);
      setJsonError('');
      return parsed;
    } catch (e) {
      setJsonError((e as Error).message);
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
      store.toast('error', (e as Error).message);
      return;
    }
    store.registerSchema(d);
    setDraft(d);
    setOriginalHadHooks(false);
    const applied = store.schema?.schema.id === d.id;
    if (applied) {
      store.set('designerOpen', false);
      store.set('panel', 'issues');
      store.toast('info', `Applied “${d.name}”.`);
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
            <Icon name="schema" /> Schemas
          </h2>
          <button className="btn icon ghost" onClick={() => store.set('designerOpen', false)} aria-label="Close">
            <Icon name="close" />
          </button>
        </header>
        <div className="designer-body">
          <nav className="designer-list">
            <h3>Registered</h3>
            <p className="muted small">Each workbook is checked against these in order; the first that fits applies.</p>
            <ul>
              {store.schemas.map((s) => (
                <li key={s.id} className={draft?.id === s.id ? 'selected' : ''}>
                  <button className="schema-item" onClick={() => startFrom(s)}>
                    <strong>{s.name}</strong>
                    <span className="mono small">{s.id}</span>
                    <span className="tags">
                      {s.id === activeId && <span className="tag tag-ok">applied</span>}
                      {hasHooks(s) && <span className="tag">code hooks</span>}
                    </span>
                  </button>
                  <div className="item-actions">
                    <button className="btn icon ghost" title="Export JSON" aria-label={`Export ${s.name}`} onClick={() => download(`${s.id}.schema.json`, schemaToJSON(s))}>
                      <Icon name="download" size={14} />
                    </button>
                    <button className="btn icon ghost" title="Remove" aria-label={`Remove ${s.name}`} onClick={() => store.removeSchema(s.id)}>
                      <Icon name="trash" size={14} />
                    </button>
                  </div>
                </li>
              ))}
            </ul>
            <div className="list-actions">
              <button className="btn primary" disabled={!ready} onClick={generate} title={ready ? '' : 'Open a workbook first'}>
                <Icon name="sparkle" /> Generate from this workbook
              </button>
              <button className="btn" onClick={() => importInput.current?.click()}>
                <Icon name="upload" /> Import JSON…
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
                <p>Pick a schema to edit, generate one from the open workbook, or import a .schema.json file.</p>
              </div>
            )}
            {draft && (
              <>
                <Tabs
                  label="Edit the schema as"
                  tabs={[
                    { id: 'form', label: 'Form' },
                    { id: 'json', label: 'JSON' },
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
                        <p className="hint warn">This schema has code hooks (async checks, business logic). They aren't part of the JSON, so applying or exporting it here leaves them out.</p>
                      )}
                    </div>
                  )}
                >
                  {mode === 'json' ? (
                    <div className="json-editor">
                      <textarea className="mono" value={jsonText} spellCheck={false} onChange={(e) => setJsonText(e.target.value)} aria-label="Schema JSON" />
                      {jsonError && <p className="hint error">{jsonError}</p>}
                    </div>
                  ) : (
                    <div className="form">
                      <div className="grid-form">
                        <label>
                          Name
                          <input value={draft.name} onChange={(e) => edit((d) => void (d.name = e.target.value))} />
                        </label>
                        <label>
                          Id
                          <input className="mono" value={draft.id} onChange={(e) => edit((d) => void (d.id = e.target.value.replace(/\s+/g, '-')))} />
                        </label>
                        <label>
                          Version
                          <input className="mono" value={draft.version ?? ''} placeholder="e.g. 2" onChange={(e) => edit((d) => void (d.version = e.target.value.trim() || undefined))} title="Saved files are stamped with it; a file saved under another version says so when opened" />
                        </label>
                        <label>
                          Hook time limit (s)
                          <input className="narrow" value={draft.hookTimeoutMs === undefined ? '' : draft.hookTimeoutMs / 1000} placeholder="10" onChange={(e) => edit((d) => {
                            const n = numberOrUndefined(e.target.value);
                            d.hookTimeoutMs = n === undefined ? undefined : Math.round(n * 1000);
                          })} />
                        </label>
                        <label className="wide">
                          Description
                          <input value={draft.description ?? ''} onChange={(e) => edit((d) => void (d.description = e.target.value || undefined))} />
                        </label>
                        <label>
                          Other sheets
                          <select value={draft.otherSheets ?? 'editable'} onChange={(e) => edit((d) => void (d.otherSheets = e.target.value as 'editable'))}>
                            <option value="editable">editable</option>
                            <option value="readOnly">read-only</option>
                          </select>
                        </label>
                        <label>
                          Saving with errors
                          <select value={draft.saveWithErrors ?? 'block'} onChange={(e) => edit((d) => void (d.saveWithErrors = e.target.value as 'block'))}>
                            <option value="block">blocked</option>
                            <option value="confirm">asks first</option>
                          </select>
                        </label>
                        <label>
                          Numbers
                          <select
                            value={draft.numbers?.display === 'full' ? 'exact' : 'formatted'}
                            onChange={(e) => edit((d) => void (d.numbers = e.target.value === 'exact' ? { display: 'full', copy: 'raw' } : undefined))}
                          >
                            <option value="formatted">as formatted</option>
                            <option value="exact">exact (every digit)</option>
                          </select>
                        </label>
                      </div>

                      {draft.sheets.map((sheet, si) => (
                        <div key={si} className="sheet-block">
                          <div className="sheet-head">
                            <label>
                              Sheet
                              <input value={matchLabel(sheet.name)} onChange={(e) => edit((d) => void (d.sheets[si].name = parseLabel(e.target.value)))} />
                            </label>
                            <label>
                              Cells outside its tables
                              <select value={sheet.outside ?? 'readOnly'} onChange={(e) => edit((d) => void (d.sheets[si].outside = e.target.value as 'readOnly'))}>
                                <option value="readOnly">read-only</option>
                                <option value="editable">editable</option>
                              </select>
                            </label>
                            <label className="inline">
                              <input type="checkbox" checked={!!sheet.optional} onChange={(e) => edit((d) => void (d.sheets[si].optional = e.target.checked || undefined))} /> optional
                            </label>
                            {!viewFor(sheet) && ready && <span className="tag tag-warn">not in this workbook</span>}
                            <button className="btn icon ghost" aria-label="Remove sheet" title="Remove sheet" onClick={() => edit((d) => void d.sheets.splice(si, 1))}>
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
                                  Table
                                  <input className="mono" value={table.key} onChange={(e) => edit((d) => void (d.sheets[si].tables![ti].key = e.target.value))} />
                                </label>
                                <label>
                                  Starts at
                                  <input className="mono narrow" value={table.at} onChange={(e) => edit((d) => void (d.sheets[si].tables![ti].at = e.target.value.toUpperCase()))} />
                                </label>
                                <label>
                                  Records run
                                  <select value={table.orientation ?? 'rows'} onChange={(e) => reread(si, ti, { orientation: e.target.value as 'rows' })}>
                                    <option value="rows">down (one per row)</option>
                                    <option value="columns">across (one per column)</option>
                                  </select>
                                </label>
                                <label className="inline" title="The first column (or row) holds fixed labels, so the records are fixed too">
                                  <input type="checkbox" checked={!!table.records?.labels} onChange={(e) => reread(si, ti, { fixed: e.target.checked })} /> fixed{' '}
                                  {table.orientation === 'columns' ? 'columns' : 'rows'}
                                </label>
                                {!table.records?.labels && (
                                  <>
                                    <label>
                                      Min
                                      <input className="narrow" value={table.records?.min ?? ''} onChange={(e) => edit((d) => {
                                        const t = d.sheets[si].tables![ti];
                                        t.records = { ...t.records, min: numberOrUndefined(e.target.value) };
                                      })} />
                                    </label>
                                    <label>
                                      Max
                                      <input className="narrow" value={table.records?.max ?? ''} onChange={(e) => edit((d) => {
                                        const t = d.sheets[si].tables![ti];
                                        t.records = { ...t.records, max: numberOrUndefined(e.target.value) };
                                      })} />
                                    </label>
                                  </>
                                )}
                                <button className="btn" disabled={!viewFor(sheet)} onClick={() => reread(si, ti, {})} title="Infer the fields again from the workbook">
                                  Re-read from workbook
                                </button>
                              </div>
                              {table.records?.labels && (
                                <p className="muted small">
                                  Fixed {table.orientation === 'columns' ? 'column' : 'row'} labels: {table.records.labels.map((l) => matchLabel(l)).join(', ')}
                                </p>
                              )}
                              <div className="fields-wrap">
                                <table className="fields">
                                  <thead>
                                    <tr>
                                      <th>Label</th>
                                      <th>Key</th>
                                      <th>Type</th>
                                      <th title="Required">Req</th>
                                      <th title="Unique">Uniq</th>
                                      <th title="Read-only">R/O</th>
                                      <th>Formulas</th>
                                      <th>Min</th>
                                      <th>Max</th>
                                      <th title="Decimal places">DP</th>
                                      <th>Values / pattern</th>
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
                              <ChecksEditor table={table} onChange={(t) => edit((d) => void (d.sheets[si].tables![ti] = t))} />
                              <button className="btn ghost small" onClick={() => edit((d) => void d.sheets[si].tables!.splice(ti, 1))}>
                                <Icon name="trash" size={14} /> Remove table
                              </button>
                            </div>
                          ))}
                        </div>
                      ))}
                      {ready && (
                        <div className="pick-actions">
                          <span className="muted small">Add from the open workbook — select cells on {store.view?.info.name ?? 'the sheet'}:</span>
                          <button className="btn" onClick={() => store.set('designerPick', 'table')}>
                            <Icon name="plus" size={14} /> Table from a selection…
                          </button>
                          <button className="btn" onClick={() => store.set('designerPick', 'cell')}>
                            <Icon name="plus" size={14} /> Single cell…
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
                      <strong>{testResult.ok ? 'Matches the open workbook.' : "Doesn't match the open workbook."}</strong>
                      {testResult.problems.length > 0 && (
                        <ul>
                          {testResult.problems.map((p) => (
                            <li key={p}>{p}</li>
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
              <Icon name="download" /> Export JSON
            </button>
            <span className="spacer" />
            <button className="btn" disabled={!ready} onClick={test}>
              Test against this workbook
            </button>
            <button className="btn primary" onClick={apply}>
              <Icon name="check" /> Apply
            </button>
          </footer>
        )}
      </div>
    </div>
  );
}

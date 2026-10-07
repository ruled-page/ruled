import { useMemo, useRef, useState } from 'react';
import { parseA1 } from '../core/refs';
import { bindSchema } from '../schema/binding';
import { generateSchema, inferTable } from '../schema/generate';
import { matchLabel, schemaFromJSON, schemaToJSON, textMatches } from '../schema/json';
import type { DeclarativeRule, FieldSchema, FieldType, SheetSchema, TableSchema, TextMatch, WorkbookSchema } from '../schema/types';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';

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
      <td>
        <input className="narrow" value={boundText(f.min, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? 'YYYY-MM-DD' : ''} onChange={(e) => set('min', boundOrUndefined(e.target.value, f.type))} aria-label="Minimum" />
      </td>
      <td>
        <input className="narrow" value={boundText(f.max, f.type)} disabled={!numeric && f.type !== 'date'} placeholder={f.type === 'date' ? 'YYYY-MM-DD' : ''} onChange={(e) => set('max', boundOrUndefined(e.target.value, f.type))} aria-label="Maximum" />
      </td>
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

function ChecksEditor({ table, onChange }: { table: TableSchema; onChange: (t: TableSchema) => void }) {
  const [kind, setKind] = useState<DeclarativeRule['rule']>('monotonic');
  const [field, setField] = useState('*');
  const [param, setParam] = useState('nonIncreasing');
  const checks = table.checks ?? [];
  const describe = (c: DeclarativeRule) => {
    const f = c.field === '*' ? 'every field' : c.field;
    if (c.rule === 'monotonic') return `${f}: ${c.direction} down the table`;
    if (c.rule === 'firstEquals') return `${f}: first value is ${c.value}`;
    return `${f}: total ${c.equals !== undefined ? `= ${c.equals}` : ''}${c.max !== undefined ? ` ≤ ${c.max}` : ''}${c.min !== undefined ? ` ≥ ${c.min}` : ''}`;
  };
  const add = () => {
    let c: DeclarativeRule;
    if (kind === 'monotonic') c = { rule: 'monotonic', field, direction: param as 'nonIncreasing' };
    else if (kind === 'firstEquals') c = { rule: 'firstEquals', field, value: Number.isNaN(Number(param)) ? param : Number(param) };
    else c = { rule: 'sum', field, max: Number(param) };
    onChange({ ...table, checks: [...checks, c] });
  };
  return (
    <div className="checks">
      <h5>Table checks</h5>
      {checks.length === 0 && <p className="muted small">None. Field rules above still apply.</p>}
      <ul>
        {checks.map((c, i) => (
          <li key={i}>
            <span>{describe(c)}</span>
            <button className="btn icon ghost" aria-label="Remove check" onClick={() => onChange({ ...table, checks: checks.filter((_, j) => j !== i) })}>
              <Icon name="trash" size={14} />
            </button>
          </li>
        ))}
      </ul>
      <div className="add-check">
        <select value={kind} onChange={(e) => {
          const k = e.target.value as DeclarativeRule['rule'];
          setKind(k);
          setParam(k === 'monotonic' ? 'nonIncreasing' : k === 'firstEquals' ? '100' : '0');
        }} aria-label="Check type">
          <option value="monotonic">Values only go one way</option>
          <option value="firstEquals">First value equals</option>
          <option value="sum">Total at most</option>
        </select>
        <select value={field} onChange={(e) => setField(e.target.value)} aria-label="Field">
          <option value="*">every field</option>
          {table.fields.map((f) => (
            <option key={f.key} value={f.key}>
              {f.key}
            </option>
          ))}
        </select>
        {kind === 'monotonic' ? (
          <select value={param} onChange={(e) => setParam(e.target.value)} aria-label="Direction">
            <option value="nonIncreasing">never go up</option>
            <option value="nonDecreasing">never go down</option>
            <option value="increasing">always go up</option>
            <option value="decreasing">always go down</option>
          </select>
        ) : (
          <input className="narrow" value={param} onChange={(e) => setParam(e.target.value)} aria-label="Value" />
        )}
        <button className="btn" onClick={add}>
          <Icon name="plus" size={14} /> Add
        </button>
      </div>
    </div>
  );
}

export function SchemaDesigner({ store }: { store: WorkbookStore }) {
  const [draft, setDraft] = useState<WorkbookSchema | null>(() => (store.schema ? toDraft(store.schema.schema) : null));
  const [originalHadHooks, setOriginalHadHooks] = useState(() => (store.schema ? hasHooks(store.schema.schema) : false));
  const [mode, setMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState('');
  const [jsonError, setJsonError] = useState('');
  const [testResult, setTestResult] = useState<{ ok: boolean; problems: string[] } | null>(null);
  const importInput = useRef<HTMLInputElement>(null);
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

  return (
    <div className="modal-backdrop" onMouseDown={() => store.set('designerOpen', false)}>
      <div className="designer" role="dialog" aria-modal="true" aria-labelledby="designer-title" onMouseDown={(e) => e.stopPropagation()}>
        <header className="designer-head">
          <h2 id="designer-title">
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
                <div className="editor-tabs">
                  <div className="panel-tabs" role="tablist">
                    <button role="tab" aria-selected={mode === 'form'} className={mode === 'form' ? 'active' : ''} onClick={() => setMode('form')}>
                      Form
                    </button>
                    <button role="tab" aria-selected={mode === 'json'} className={mode === 'json' ? 'active' : ''} onClick={() => {
                      setJsonText(json);
                      setJsonError('');
                      setMode('json');
                    }}>
                      JSON
                    </button>
                  </div>
                  {originalHadHooks && (
                    <p className="hint warn">This schema has code hooks (async checks, business logic). They aren't part of the JSON, so applying or exporting it here leaves them out.</p>
                  )}
                </div>

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
                        {(sheet.cells?.length ?? 0) > 0 && (
                          <p className="muted small">
                            Also checks {sheet.cells!.length} single cell{sheet.cells!.length === 1 ? '' : 's'} ({sheet.cells!.map((c) => c.ref).join(', ')}); edit those in the JSON view.
                          </p>
                        )}
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
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                )}

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

import { ChangeDetectionStrategy, Component, CUSTOM_ELEMENTS_SCHEMA, ElementRef, signal, viewChild } from '@angular/core';
import { DecimalPipe, PercentPipe } from '@angular/common';
import type { ChangeEventDetail, EditorOptions, SaveEventDetail, WorkbookEditorElement, WorkbookSchema } from '@ruled/sheet/standalone';

// The editor is a few megabytes (React, the grid and its worker are inside),
// so it loads after Angular has drawn the page. <workbook-editor> upgrades
// when it arrives, keeping the options Angular has already set on it.
const editorLoaded = import('@ruled/sheet/standalone');

/** The forecast workbook's rules, as Angular's side of the app defines them. */
const forecastSchema: WorkbookSchema = {
  id: 'example.forecast',
  name: { en: 'Forecast (from Angular)', de: 'Prognose (aus Angular)' },
  version: '1',
  sheets: [
    {
      name: 'Forecast',
      tables: [
        {
          key: 'quarters',
          at: 'A1',
          orientation: 'columns',
          fields: [
            { key: 'period', label: 'Period', type: 'text', required: true, pattern: '^\\d{4}-Q[1-4]$' },
            { key: 'revenue', label: 'Revenue', type: 'number', min: 0 },
            { key: 'cost', label: 'Cost', type: 'number', min: 0 },
            { key: 'margin', label: 'Margin', type: 'number', formula: 'require' },
            { key: 'marginPct', label: 'Margin %', type: 'percent', formula: 'require' },
          ],
        },
      ],
    },
  ],
};

interface Quarter {
  period: string;
  revenue: number;
  cost: number;
  margin: number;
  marginPct: number;
}

interface Received {
  fileName: string;
  bytes: number;
  quarters: Quarter[];
  changes: string[];
}

@Component({
  selector: 'app-root',
  // <workbook-editor> is a custom element, not an Angular component.
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  imports: [DecimalPipe, PercentPipe],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <header>
      <h1>Forecast review</h1>
      <span class="status" role="status">{{ status() }}</span>
      <label>
        Language
        <select (change)="setLanguage($any($event.target).value)">
          <option value="en">English</option>
          <option value="de" lang="de">Deutsch</option>
        </select>
      </label>
      <button type="button" (click)="save()">Save from Angular</button>
    </header>

    <workbook-editor
      #editor
      save-handler="event"
      src="samples/forecast.xlsx"
      [options]="options"
      (change)="onChange($event)"
      (save)="onSave($event)"
    ></workbook-editor>

    <section aria-labelledby="received-heading">
      <h2 id="received-heading">What Angular received</h2>
      @if (received(); as r) {
        <p>
          <strong>{{ r.fileName }}</strong>, {{ r.bytes | number }} bytes. Changed: {{ r.changes.join(', ') || 'nothing' }}.
        </p>
        <table>
          <thead>
            <tr>
              <th scope="col">Period</th>
              <th scope="col">Revenue</th>
              <th scope="col">Cost</th>
              <th scope="col">Margin</th>
              <th scope="col">Margin %</th>
            </tr>
          </thead>
          <tbody>
            @for (q of r.quarters; track q.period) {
              <tr>
                <th scope="row">{{ q.period }}</th>
                <td>{{ q.revenue | number: '1.0-2' }}</td>
                <td>{{ q.cost | number: '1.0-2' }}</td>
                <td>{{ q.margin | number: '1.0-2' }}</td>
                <td>{{ q.marginPct | percent: '1.0-1' }}</td>
              </tr>
            }
          </tbody>
        </table>
      } @else {
        <p class="muted">Nothing saved yet. Edit the forecast and press Save.</p>
      }
    </section>
  `,
  styles: `
    :host {
      display: grid;
      grid-template-rows: auto 1fr auto;
      height: 100vh;
    }
    header {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 12px 20px;
      padding: 10px 20px;
      background: #fff;
      border-bottom: 1px solid var(--host-border);
    }
    h1 {
      font-size: 18px;
      margin: 0;
      flex: 1;
    }
    .status {
      color: var(--host-muted);
    }
    button,
    select {
      font: inherit;
      padding: 6px 12px;
      border-radius: 6px;
      border: 1px solid var(--host-border);
      background: #fff;
      color: var(--host-text);
    }
    button {
      background: var(--host-accent);
      border-color: var(--host-accent);
      color: #fff;
    }
    workbook-editor {
      min-height: 0;
    }
    section {
      max-height: 34vh;
      overflow: auto;
      padding: 8px 20px 16px;
      background: #fff;
      border-top: 1px solid var(--host-border);
    }
    h2 {
      font-size: 15px;
      margin: 8px 0;
    }
    table {
      border-collapse: collapse;
    }
    th,
    td {
      padding: 4px 14px 4px 0;
      text-align: end;
      font-variant-numeric: tabular-nums;
    }
    th[scope='row'],
    thead th:first-child {
      text-align: start;
    }
    .muted {
      color: var(--host-muted);
    }
  `,
})
export class App {
  private readonly editor = viewChild.required<ElementRef<WorkbookEditorElement>>('editor');

  protected readonly options: EditorOptions = {
    schemas: [forecastSchema],
    saveLabel: { en: 'Save to Angular', de: 'In Angular speichern' },
    features: { open: false },
    warnOnUnload: false,
  };

  protected readonly status = signal('No changes');
  protected readonly received = signal<Received | undefined>(undefined);

  protected onChange(event: Event): void {
    const { editedCount: n, unsaved } = (event as CustomEvent<ChangeEventDetail>).detail;
    this.status.set(n ? `${n} change${n === 1 ? '' : 's'}${unsaved ? ', not saved' : ', saved'}` : 'No changes');
  }

  /** The editor's save, handed to Angular: keep the file and read the values. */
  protected onSave(event: Event): void {
    const saved = (event as CustomEvent<SaveEventDetail>).detail;
    const quarters = (saved.values?.['Forecast']?.tables['quarters'] ?? []) as unknown as Quarter[];
    // A real app would upload saved.bytes or saved.values here; the save waits for it.
    saved.waitUntil(new Promise((resolve) => setTimeout(resolve, 200)));
    this.received.set({
      fileName: saved.fileName,
      bytes: saved.bytes.length,
      quarters,
      changes: saved.changes.map((c) => `${c.sheet}!${c.ref}`),
    });
  }

  protected async save(): Promise<void> {
    await editorLoaded;
    await this.editor().nativeElement.save();
  }

  protected async setLanguage(language: string): Promise<void> {
    await editorLoaded;
    this.editor().nativeElement.setLanguage(language);
  }
}

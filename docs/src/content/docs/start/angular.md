---
title: Angular
description: Use the editor's element in an Angular app, loaded only when it's needed.
sidebar:
  order: 6
---

In Angular the editor is a custom element like any other. Allow it with `CUSTOM_ELEMENTS_SCHEMA`, bind its `options`, and listen to its events. Loading the standalone build with a dynamic `import()` keeps it out of your initial bundle; the element keeps any options Angular sets before it has loaded.

```ts
import { Component, CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';
import type { EditorOptions, SaveEventDetail } from '@sitepen/spreadsheet/standalone';
import { forecast } from './schemas';

const editorLoaded = import('@sitepen/spreadsheet/standalone');

@Component({
  selector: 'app-forecast-review',
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  template: `
    <workbook-editor
      save-handler="event"
      [attr.src]="fileUrl"
      [options]="options"
      (save)="onSave($event)"
    ></workbook-editor>
  `,
  styles: [`workbook-editor { display: block; height: 80vh; }`],
})
export class ForecastReview {
  fileUrl = '/files/forecast.xlsx';
  options: EditorOptions = { schemas: [forecast], saveLabel: 'Save to the app', features: { open: false } };

  constructor(private api: ForecastApi) {}

  onSave(event: Event) {
    const saved = (event as CustomEvent<SaveEventDetail>).detail;
    saved.waitUntil(this.api.upload(saved.bytes, saved.values));
  }
}
```

The standalone build has type declarations, including `WorkbookEditorElement` for the element itself, so `@ViewChild` gives you a typed element to call `goTo`, `save` or `getWorkbook` on.

The [web component page](/start/web-component/) lists the element's attributes, methods and events.

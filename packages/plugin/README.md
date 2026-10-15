# @ruled/plugin

**Ruled's own plugins** for its spreadsheet editor,
[`@ruled/sheet`](https://www.npmjs.com/package/@ruled/sheet): open workbooks
from cloud drives, save them back, and save copies there. Each connector is a
subpath, so a host bundles only the ones it imports.

- **Documentation:** [ruled.page](https://ruled.page/plugins/onedrive/)
- **Source:** [github.com/ruled-page/ruled](https://github.com/ruled-page/ruled)

## Install

```sh
npm install @ruled/sheet @ruled/plugin
```

## OneDrive and SharePoint

```ts
import { useWorkbookEditor } from '@ruled/sheet';
import { oneDrive } from '@ruled/plugin/onedrive';

const editor = useWorkbookEditor({
  plugins: [oneDrive({ clientId: 'your-entra-app-id' })],
});
```

The editor's Open menu gains OneDrive, a workbook opened from it can be saved
back with "Save to OneDrive", and "Save a copy" can write a new file there.
Sign-in uses Microsoft's popups, so the page also needs a small return page
that runs `@ruled/plugin/onedrive/redirect`. The documentation walks through
registering the app in Microsoft Entra ID, that page, and the Content
Security Policy.

## Licence

MIT

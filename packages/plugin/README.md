# @ruled/plugin

**Ruled's own plugins** for its spreadsheet editor,
[`@ruled/sheet`](https://www.npmjs.com/package/@ruled/sheet): open workbooks
from cloud drives, save them back, and save copies there. Each connector is a
subpath, so a host bundles only the ones it imports.

- **Documentation:** [OneDrive](https://ruled.page/plugins/onedrive/), [Google Drive](https://ruled.page/plugins/google-drive/)
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

## Google Drive

```ts
import { useWorkbookEditor } from '@ruled/sheet';
import { googleDrive } from '@ruled/plugin/google-drive';

const editor = useWorkbookEditor({
  plugins: [googleDrive({ clientId: 'your-client-id.apps.googleusercontent.com', apiKey: 'your-api-key', appId: 'your-project-number' })],
});
```

The Open menu gains Google Drive. An `.xlsx` opened from it can be saved
back; a Google Sheet opens as an Excel copy, which "Save a copy to Google
Drive" writes as a new `.xlsx`, leaving the Google Sheet as it is. Sign-in
asks only for access to the files people choose (`drive.file`). The
documentation walks through setting up the Google Cloud project and the
Content Security Policy.

## Licence

MIT

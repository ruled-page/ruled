---
title: OneDrive and SharePoint
description: Open workbooks from OneDrive and SharePoint, save them back, and save copies there, with the @ruled/plugin/onedrive connector.
sidebar:
  order: 1
---

`@ruled/plugin/onedrive` connects the editor to Microsoft's cloud storage: a person's own OneDrive (personal, or work and school), and SharePoint document libraries. With it:

- the Open button becomes a menu with **OneDrive** beside **This device**, which opens Microsoft's own file picker;
- a workbook opened from OneDrive can be written back with **Save to OneDrive**, which refuses to overwrite the file if someone else saved it in the meantime;
- **Save a copy** can write a new file into a OneDrive folder.

Everything runs in the browser: the person signs in with Microsoft in a popup, and the editor talks to Microsoft Graph directly. Your server never sees the file or a token. The file is `.xlsx` at both ends, so nothing is converted.

Plugins are trusted code running on your page, like any other dependency.

## Install

```sh
npm install @ruled/sheet @ruled/plugin
```

`@ruled/plugin` brings Microsoft's sign-in library, MSAL (`@azure/msal-browser`), with it. Your bundle loads MSAL only when someone first opens the OneDrive menu, so the editor itself is no larger.

## Use it

Pass the connector in `plugins`, with your app registration's client id (see [Register an app](#register-an-app-in-microsoft-entra-id) below):

```tsx
import { useWorkbookEditor, WorkbookEditor } from '@ruled/sheet';
import { oneDrive } from '@ruled/plugin/onedrive';

function Prices() {
  const editor = useWorkbookEditor({
    plugins: [oneDrive({ clientId: '11111111-2222-3333-4444-555555555555' })],
  });
  return <WorkbookEditor editor={editor} />;
}
```

The `<workbook-editor>` element takes it through its `options` property:

```js
import { oneDrive } from '@ruled/plugin/onedrive';

document.querySelector('workbook-editor').options = {
  plugins: [oneDrive({ clientId: '11111111-2222-3333-4444-555555555555' })],
};
```

`plugins` is read when the editor is made. To change it, make a new editor (a new `key` in React, or a new element).

## The sign-in return page

Microsoft's sign-in popup ends on a page of yours, which hands the result back to the editor's page and closes. That page runs one module from the connector and nothing else:

```html
<!-- onedrive-redirect.html, beside your app's index.html -->
<!doctype html>
<html lang="en">
  <head><meta charset="UTF-8" /><title>Signing in…</title></head>
  <body>
    <script type="module">
      import '@ruled/plugin/onedrive/redirect';
    </script>
  </body>
</html>
```

Bundle it with your app, so it's served from your origin. With Vite, add it as a second input:

```ts
// vite.config.ts
export default defineConfig({
  build: { rollupOptions: { input: ['index.html', 'onedrive-redirect.html'] } },
});
```

By default the connector expects the page at `/onedrive-redirect.html` on the editor page's origin. Put it elsewhere with `redirectUri`. Wherever it is, register its exact address in your app registration, and serve it:

- **without** a `Cross-Origin-Opener-Policy` header, which would stop it reaching the page that opened it;
- with `Cache-Control: no-store`, since the sign-in's answer passes through it.

## Register an app in Microsoft Entra ID

The connector signs people in as your app, which you register once with Microsoft. You need a Microsoft Entra tenant to register it in: your organisation's, or a free one.

1. In the [Microsoft Entra admin center](https://entra.microsoft.com), go to **Identity › Applications › App registrations** and choose **New registration**.
2. Give it a name people will recognise: they see it when they sign in.
3. Under **Supported account types**, choose who may sign in. For both personal accounts and work or school ones, choose **Accounts in any organizational directory and personal Microsoft accounts**. That matches the connector's default `authority`, `'common'`.
4. Under **Redirect URI**, choose the **Single-page application (SPA)** platform and enter your return page's address, such as `https://app.example.com/onedrive-redirect.html`. Then **Register**.
5. Copy the **Application (client) ID** from the overview. That's `clientId`.
6. Under **API permissions**, choose **Add a permission**, then:
   - **Microsoft Graph › Delegated permissions:** `Files.ReadWrite.All`, so the connector can open and save the files people pick, including ones shared with them and in SharePoint libraries. (`Files.Read.All` with `readOnly: true`.) Keep `User.Read`.
   - **SharePoint › Delegated permissions**, for work and school accounts, whose file picker runs on SharePoint: `MyFiles.Write` and `AllSites.Write` (`MyFiles.Read` and `AllSites.Read` with `readOnly: true`).
7. Some organisations only let an administrator consent to these. If people see "Need admin approval" when they sign in, an administrator chooses **Grant admin consent for …** on the same page.

Every address the editor runs on needs its own redirect URI under the SPA platform, the local one you develop on included (see [Try it locally](#try-it-locally)).

## Options

```ts
oneDrive({
  clientId: string,
  authority?: 'common' | 'organizations' | 'consumers' | string,
  redirectUri?: string,
  sharePoint?: { baseUrl: string },
  readOnly?: boolean,
  label?: Localized,
});
```

| Option | Default | What it does |
| --- | --- | --- |
| `clientId` | (required) | Your app registration's application (client) id. |
| `authority` | `'common'` | Who can sign in: `'common'` (personal, work and school accounts), `'organizations'` (work and school only), `'consumers'` (personal only), or a tenant's id or domain, such as `'contoso.com'`, for one organisation. It must agree with the account types your app registration supports. |
| `redirectUri` | `/onedrive-redirect.html` on the page's origin | The [return page](#the-sign-in-return-page). |
| `sharePoint` | none | `{ baseUrl: 'https://contoso.sharepoint.com/sites/finance' }` opens the picker in that SharePoint site rather than the person's OneDrive. For work and school accounts only. |
| `readOnly` | `false` | Asks only for read permission. Workbooks open, but there's no **Save to OneDrive** and no copies to OneDrive. |
| `label` | "OneDrive", or "SharePoint" with `sharePoint` | The name in the Open menu, on **Save to …** and in **Save a copy**, as written text or `{ en: '…', de: '…' }`. |

## What it does

**Signing in.** The first time someone opens the OneDrive menu, a Microsoft sign-in popup opens. The sign-in lasts as long as the browser tab: MSAL keeps it in `sessionStorage`, never with the workbook or its drafts. **Sign out of OneDrive**, in the More menu, forgets it in this editor without signing the person out of Microsoft elsewhere.

Browsers allow a popup only straight after a click. When signing in has used up that moment, the editor asks once more ("Choose a workbook in OneDrive") and opens the picker on that click.

**Opening.** The picker shows the files the editor opens: `.xlsx`, `.xlsm`, `.csv` and `.tsv`. The connector reads the file's details first, and refuses one larger than the editor's `limits.maxFileBytes` (50 MB unless you change it) before downloading anything. The file's folder is shown when the pointer rests on its name.

**Saving back.** **Save to OneDrive** writes the workbook over the file it came from, but only if the file is still the version that was opened. If someone saved it since, the editor says so and offers to download your version, as it does for any [save conflict](/guides/opening-and-saving/). Each successful save becomes the new version to check against. Files over 4 MB are uploaded in parts, and a failed upload leaves the file as it was.

**Saving a copy.** **Save a copy › OneDrive** opens the picker on folders. The copy takes the workbook's name. If the folder already has a file by that name, the connector asks for another rather than overwriting it. The workbook open in the editor stays the one it came from.

How these sit beside your app's own Save is covered in [Opening and saving](/guides/opening-and-saving/): Save to OneDrive is separate from your `onSave`, and `features.sourceSave: false` turns writing to OneDrive off.

## Content Security Policy

Beside what the editor itself needs ([Install › Content Security Policy](/start/install/#content-security-policy)), allow:

| Directive | Add |
| --- | --- |
| `connect-src` | `https://login.microsoftonline.com https://graph.microsoft.com`, and the hosts files are downloaded from and uploaded to: `https://*.sharepoint.com` for work and school accounts, and `https://*.1drv.com https://*.microsoftpersonalcontent.com https://api.onedrive.com` for personal ones |
| `frame-src` | `https://login.microsoftonline.com`, where MSAL renews a sign-in in a hidden frame |
| `script-src` | nothing: MSAL is bundled with your app |

The page with the editor must not send `Cross-Origin-Opener-Policy: same-origin`, which cuts it off from the sign-in and picker popups. Send `same-origin-allow-popups`, or no header. The [return page](#the-sign-in-return-page) must send none.

## Language

The connector's text is in English and German, under `plugin.page.ruled.onedrive.<key>`. Your `translations` can reword it, or add a language, like any of the editor's messages:

```ts
useWorkbookEditor({
  plugins: [oneDrive({ clientId, label: { en: 'Company files', de: 'Firmendateien' } })],
  translations: { fr: { 'plugin.page.ruled.onedrive.signIn': 'Se connecter' } },
});
```

The picker itself is Microsoft's. It's in German when the editor is, and otherwise in English.

## Try it locally

Microsoft accepts `http://localhost` redirect addresses, so you can try the connector against your own accounts before deploying. In Ruled's own repository, the demo's dev server loads the connector when it's given a client id:

1. [Register an app](#register-an-app-in-microsoft-entra-id) with the redirect URI `http://localhost:5180/onedrive-redirect.html` under the **Single-page application** platform.
2. Create `examples/demo/.env.local`, which git ignores:

   ```sh
   VITE_ONEDRIVE_CLIENT_ID=11111111-2222-3333-4444-555555555555
   # Optional: 'organizations', 'consumers' or a tenant; 'common' when left out.
   VITE_ONEDRIVE_AUTHORITY=common
   ```

3. Run `npm run dev` and open `http://localhost:5180`. **Open** now has a OneDrive entry.

The demo's return page is `examples/demo/onedrive-redirect.html`. Built or deployed, the demo has no connector: it reads the client id only in development.

In your own app, the same applies: register `http://localhost:<port>/<your return page>` beside your real addresses, and use a separate app registration for production if you'd rather keep them apart.

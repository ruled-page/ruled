---
title: Google Drive
description: Open workbooks and Google Sheets from Google Drive, save workbooks back, and save copies there, with the @ruled/plugin/google-drive connector.
sidebar:
  order: 4
---

`@ruled/plugin/google-drive` connects the editor to Google Drive. With it:

- the Open button becomes a menu with **Google Drive** beside **This device**, which opens Google's own file picker;
- an Excel workbook opened from Drive can be written back with **Save to Google Drive**, which refuses to overwrite the file if someone else changed it in the meantime;
- a Google Sheet opens as an Excel copy, which **Save a copy to Google Drive** writes as a new `.xlsx`; the Google Sheet itself is never written over;
- **Save a copy** can write a new file into a Drive folder.

Everything runs in the browser: the person signs in with Google in a popup, and the editor talks to the Drive API directly. Your server never sees the file or a token.

Plugins are [trusted code](/plugins/using/#plugins-are-trusted-code) running on your page, like any other dependency.

## Install

```sh
npm install @ruled/sheet @ruled/plugin
```

Google's sign-in library and its Picker aren't on npm. The connector loads them from Google the first time someone opens the Google Drive menu, along with the rest of its own code, so your bundle carries only a small entry, about 1 kB gzipped.

## Use it

Pass the connector in `plugins`, with three values from your Google Cloud project (see [Set up a Google Cloud project](#set-up-a-google-cloud-project) below):

```tsx
import { useWorkbookEditor, WorkbookEditor } from '@ruled/sheet';
import { googleDrive } from '@ruled/plugin/google-drive';

function Prices() {
  const editor = useWorkbookEditor({
    plugins: [
      googleDrive({
        clientId: '123456789012-abc.apps.googleusercontent.com',
        apiKey: 'AIzaSy…',
        appId: '123456789012',
      }),
    ],
  });
  return <WorkbookEditor editor={editor} />;
}
```

The `<workbook-editor>` element takes it through its `options` property:

```js
import { googleDrive } from '@ruled/plugin/google-drive';

document.querySelector('workbook-editor').options = {
  plugins: [googleDrive({ clientId: '123456789012-abc.apps.googleusercontent.com', apiKey: 'AIzaSy…', appId: '123456789012' })],
};
```

`plugins` is read when the editor is made. To change it, make a new editor (a new `key` in React, or a new element).

None of the three values is a secret: they identify your project to Google, and they end up in your page's code. What keeps them yours is the list of sites each is restricted to.

## Set up a Google Cloud project

The connector signs people in as your app, which you set up once in the [Google Cloud console](https://console.cloud.google.com). It asks for one permission, `drive.file`: access to the files people choose in the Picker and the files the editor creates, and nothing else in their Drive. Google counts it as non-sensitive, so your app needs no security assessment, only Google's basic checks of your app's name and sites when you publish it.

1. **Create a project**, or choose one. Its **project number**, on the project's dashboard or under **IAM & Admin › Settings**, is `appId`. (It's a number, not the project's id.)
2. **Turn on the APIs.** Under **APIs & Services › Library**, enable the **Google Drive API** and the **Google Picker API**.
3. **Describe your app** under **Google Auth Platform** (called **OAuth consent screen** in older consoles):
   - **Branding:** your app's name, which people see when they sign in, a support email, and your app's home page and privacy policy before you publish it. The privacy policy must say what your app does with Google data and that it follows Google's Limited Use rules; [Ruled's own](/privacy/) is an example.
   - **Audience:** **External**, so any Google account can sign in, or **Internal** for your Google Workspace organisation only. An external app starts in **Testing**: only the test users you list there can sign in. **Publish** it when your real site is ready.
   - **Data access:** add the scope `https://www.googleapis.com/auth/drive.file` ("See, edit, create and delete only the specific Google Drive files you use with this app").
4. **Make an OAuth client.** Under **Clients** (or **APIs & Services › Credentials › Create credentials › OAuth client ID**), choose **Web application**. Under **Authorized JavaScript origins**, add each origin the editor runs on, such as `https://app.example.com`, and `http://localhost:5180` for trying it locally. No redirect URIs are needed: sign-in happens in a popup that reports back to the page. Its **client ID** is `clientId`.
5. **Make an API key** under **APIs & Services › Credentials › Create credentials › API key**, then restrict it:
   - **Application restrictions:** **Websites**, with your sites as patterns such as `https://app.example.com/*` and `http://localhost:5180/*`.
   - **API restrictions:** **Google Picker API** only. Drive is called with the person's sign-in, not the key.

   That key is `apiKey`.

Origins are exact: `http://localhost:5180` and `http://localhost:5173` are different, and so are `https://example.com` and `https://www.example.com`. A new origin can take a few minutes to start working.

## Options

```ts
googleDrive({
  clientId: string,
  apiKey: string,
  appId: string,
  label?: Localized,
});
```

| Option | Default | What it does |
| --- | --- | --- |
| `clientId` | (required) | Your web OAuth client's id, ending `.apps.googleusercontent.com`. |
| `apiKey` | (required) | Your API key, restricted to your sites and the Picker API. The Picker needs it. |
| `appId` | (required) | Your project's number. The Picker needs it for `drive.file` to cover the files people choose. |
| `label` | "Google Drive" | The name in the Open menu, on **Save to …** and in **Save a copy**, as written text or `{ en: '…', de: '…' }`. |

## What it does

**Signing in.** The first time someone opens a file from Google Drive, Google's sign-in popup opens and asks them to let your app see the files they choose. The access lasts about an hour and is kept only in the page's memory: never in storage, and never with the workbook or its drafts. After an hour, or after a reload, the next action asks again in a popup, which usually closes by itself. **Sign out of Google Drive**, in the More menu, revokes the access the editor was given.

Browsers allow a popup only straight after a click. When the editor couldn't open Google's popup inside it, because Google's scripts were still loading, it asks once more and opens the popup on that click.

**Opening.** The Picker shows Excel workbooks (`.xlsx` and `.xlsm`) and Google Sheets, including those in shared drives. The connector reads the file's details first, and refuses a workbook larger than the editor's `limits.maxFileBytes` (50 MB unless you change it) before downloading it.

**Saving back.** **Save to Google Drive** writes the workbook over the `.xlsx` it came from, but only if the file is still the version that was opened. Drive has no way to make that check part of the upload, so the connector reads the file's version and time of last change just before writing, and if someone changed it since, the editor says so and offers to download your version, as it does for any [save conflict](/guides/opening-and-saving/). Each successful save becomes the new version to check against.

**Saving a copy.** **Save a copy › Google Drive** opens the Picker on folders, starting in the folder of the file the workbook came from. The copy takes the workbook's name. Drive allows two files with the same name in a folder, so a copy never replaces another file. The workbook open in the editor stays the one it came from.

How these sit beside your app's own Save is covered in [Using plugins](/plugins/using/#saving-with-a-source): Save to Google Drive is separate from your `onSave`, and `features.sourceSave: false` turns writing to Drive off.

## Google Sheets

A Google Sheet is in Google's own format, not Excel's. The connector asks Google to export it as an `.xlsx`, as **File › Download › Microsoft Excel** does in Google Sheets, and opens that copy. A banner says so as it opens: it's a Google Sheet, open as an Excel copy; saving makes a new `.xlsx` in Drive; and the Google Sheet stays as it is.

There's no **Save to Google Drive** for a Google Sheet. The save button reads **Save a copy to Google Drive**, and writes a new `.xlsx`. Writing over the Google Sheet would have Google convert the workbook back into its own format, losing everything the export dropped, in the file other people share. The editor never does that.

What the export loses:

- **Functions only Google Sheets has**, such as `QUERY`, `IMPORTRANGE`, `GOOGLEFINANCE`, `ARRAYFORMULA`, `SPARKLINE`, `SPLIT` and `REGEXMATCH`. Google writes each such cell as `IFERROR(__xludf.DUMMYFUNCTION("…"), <its last value>)`, so it keeps the value Google last calculated, and stops recalculating. The banner names the functions and how many cells used each.
- **Apps Script**, and features only Google Sheets has, such as smart chips, connected sheets and filter views.

Google exports Sheets of up to 10 MB. For a larger one, the editor says it's too large to open; splitting it into smaller Sheets in Google Sheets gets around that.

## Content Security Policy

Beside what the editor itself needs ([Install › Content Security Policy](/start/install/#content-security-policy)), allow:

| Directive | Add |
| --- | --- |
| `script-src` | `https://accounts.google.com/gsi/client https://apis.google.com`, for Google's sign-in library and the Picker, which loads further scripts from `apis.google.com` |
| `frame-src` | `https://docs.google.com`, where the Picker runs, in a frame over the page |
| `connect-src` | `https://www.googleapis.com`, the Drive API, which uploads go to as well |
| `style-src` | `'unsafe-inline'`: the Picker adds its styles to the page in `<style>` elements |

Signing in happens in a popup on `accounts.google.com`, which needs no entry. The page with the editor must not send `Cross-Origin-Opener-Policy: same-origin`, which cuts it off from that popup. Send `same-origin-allow-popups`, or no header.

## Language

The connector's text is in English and German, under `plugin.page.ruled.google-drive.<key>`. Your `translations` can reword it, or add a language, like any of the editor's messages:

```ts
useWorkbookEditor({
  plugins: [googleDrive({ clientId, apiKey, appId, label: { en: 'Company Drive', de: 'Firmenablage' } })],
  translations: { fr: { 'plugin.page.ruled.google-drive.signIn': 'Se connecter' } },
});
```

The Picker and the sign-in popup are Google's. The Picker is in German when the editor is, and otherwise in English; the popup follows the person's Google account.

## Try it locally

Google accepts `http://localhost` origins, so you can try the connector against your own Google account before deploying. In Ruled's own repository, the demo's dev server loads the connector when it's given your project's values:

1. [Set up a Google Cloud project](#set-up-a-google-cloud-project) with `http://localhost:5180` as an authorised JavaScript origin, `http://localhost:5180/*` on the API key, and your Google account as a test user.
2. Create `examples/demo/.env.local`, which git ignores:

   ```sh
   VITE_GOOGLE_CLIENT_ID=123456789012-abc.apps.googleusercontent.com
   VITE_GOOGLE_API_KEY=AIzaSy…
   VITE_GOOGLE_APP_ID=123456789012
   ```

3. Run `npm run dev` and open `http://localhost:5180`. **Open** now has a Google Drive entry.

Built or deployed, the demo has no connector: it reads these values only in development.

In your own app, the same applies: add `http://localhost:<port>` beside your real origins, or use a separate Google Cloud project for development if you'd rather keep them apart.

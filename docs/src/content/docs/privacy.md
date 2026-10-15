---
title: Privacy
description: What Ruled's websites and its demo do with your data, including files you open from Google Drive.
---

_Last updated 15 October 2026._

This covers the websites at ruled.page: this documentation, the demo at [demo.ruled.page](https://demo.ruled.page), and the live examples linked from them. An app that embeds Ruled's editor has its own privacy policy, which applies there instead.

## Your workbooks stay in your browser

Ruled has no server. The editor opens, calculates and saves workbooks entirely in your browser. A file you open is never sent to Ruled, and Ruled has no account, database or storage where your files or edits could be kept.

## What the sites keep in your browser

The editor and the demo keep a few things in your browser's own storage, on your device only, never sent anywhere:

- **Drafts** of unsaved edits, in IndexedDB (`workbook-editor-drafts`), so a workbook you were editing can be offered back if the page closes. They are kept against the file they belong to, and deleted when you save or discard them.
- **Your choices**, in `localStorage`: the language and theme you picked, and small values for connectors, such as recent files, under `ruled-plugin:<name>`.

You can remove all of it by clearing the site's data in your browser. The sites set no cookies, and use no analytics, advertising or tracking.

## Google Drive

The demo can open workbooks from your Google Drive and save them there, when you choose to connect it.

- **What it can reach.** When you sign in, Google asks you to let the demo "see, edit, create and delete only the specific Google Drive files you use with this app" (the `drive.file` permission). It can reach only the files you choose in Google's file picker and the copies it saves for you, never the rest of your Drive.
- **How it's used.** Your browser talks to Google directly. The demo downloads the file you chose so you can edit it, and uploads your workbook when you save it to Drive or save a copy there. It reads a file's name, type, size, folder and version, to open it and to avoid overwriting changes someone else made since you opened it. It does nothing else with your Google data.
- **What's kept.** Google gives the page an access token that lasts about an hour. It is held only in the page's memory: never stored, never sent to Ruled, and gone when you close or reload the page. Your files and their contents are not stored or shared by Ruled.
- **Stopping it.** **Sign out of Google Drive**, in the editor's More menu, revokes the access. You can also remove it at any time in your Google Account, under [Third-party apps and services](https://myaccount.google.com/connections).

Ruled's use and transfer of information received from Google APIs adheres to the [Google API Services User Data Policy](https://developers.google.com/terms/api-services-user-data-policy), including the Limited Use requirements. It is not used for advertising, sold, or used to train AI models, and no person reads it.

## OneDrive and SharePoint

When the demo offers OneDrive, the same applies: your browser talks to Microsoft directly, only to open and save the files you choose, and Microsoft's sign-in is kept in the browser tab's session storage until you close the tab or choose **Sign out of OneDrive**.

## Hosting

The sites are served by [Cloudflare Pages](https://www.cloudflare.com/privacypolicy/). Like any web host, Cloudflare handles the requests your browser makes to load the pages, which include your IP address, and may keep logs of them for security and operation under its own privacy policy. Ruled doesn't receive or use them.

## Changes and questions

If this changes, this page will say so, with a new date above. For questions, open an issue at [github.com/ruled-page/ruled](https://github.com/ruled-page/ruled/issues).

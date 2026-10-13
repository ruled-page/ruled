# Ruled

**Ruled** is a spreadsheet editor for web apps that knows the rules: it opens
real Excel workbooks, checks every cell against a schema with the app's own
sync and async checks, and saves the `.xlsx` back intact.

The name holds both halves of what it is. Ruled paper is the ledger a
spreadsheet grew from; rules are what a schema holds a workbook to.

- Packages: `@ruled/sheet` (the editor), and `@ruled/formula` for the
  calculation engine when it is released on its own.
- Line: **Spreadsheets that know the rules.**

## The mark

A page of a ledger: three ruled lines, the red margin line ruled paper has,
and an R written on the bottom line, as if by hand. The margin line's red is
the red of a validation marker in a cell's corner, so it stands for the rules
too.

| File | What it is |
| --- | --- |
| `ruled-mark.svg` | The mark in its colours, for light backgrounds |
| `ruled-mark-dark.svg` | The mark for dark backgrounds |
| `ruled-mark-mono.svg` | The mark in one colour, for one-colour printing and stamps |
| `ruled-favicon.svg` | The mark for browser tabs |
| `ruled-icon-512.png`, `ruled-apple-touch-icon.png` | The mark as app icons |
| `ruled-wordmark.svg`, `ruled-wordmark-dark.svg` | "Ruled" on its rule, crossed by the margin line |
| `ruled-lockup.svg`, `ruled-lockup-dark.svg` | The mark and the wordmark together: the usual logo |
| `ruled-social.png` | The card shown when a link to the docs is shared (1200 × 630) |

Use the lockup where there is room and the mark where there isn't. Leave
space around either at least as wide as the mark's margin line is from its
edge, and don't show the mark smaller than 16 pixels. Keep its colours: the
margin line is always red, never the brand green.

## Colour

| Name | Hex | Used for |
| --- | --- | --- |
| Ledger green | `#0B6B4F` | The mark's page; what people act on, in light themes |
| Mint | `#3ECF9A` | The same, on dark backgrounds |
| Ink | `#101A16` | Text |
| Paper | `#FBF9F4` | Backgrounds, and the mark's lines and letter |
| Night | `#0E1411` | Dark backgrounds |
| Margin red | `#E0412A`, or `#FF6B4F` on ledger green | The margin line, and marking what's checked; sparingly |
| Highlighter | `#F4D54A` | Rare emphasis, as a highlighter would |

Greys lean green, toward the ink, rather than neutral.

## Type

- **Fraunces** for the wordmark and headings: a soft, old-style serif with
  optical sizes, set heavy (600 to 700) with tight spacing. Italic for
  emphasis in headlines.
- **IBM Plex Sans** for text, and **IBM Plex Mono** for code, cell
  references and small labels.

All three are free (SIL Open Font License); the docs serve them from
`@fontsource`.

## Remaking the assets

The letters in the SVG files are Fraunces' own outlines, so the files need no
fonts to show. `tools/` makes them:

1. `tools/outline.py` outlines a word from the Fraunces variable font at
   given settings into `tools/ruled.json` and `tools/R.json` (needs
   `pip install fonttools brotli`, and the docs' dependencies installed).
2. `python3 brand/tools/make.py brand` writes the SVG files.
3. `node brand/tools/render.mjs` renders the social card (`tools/social.html`)
   and the icons as PNG files.

The docs copy what they use from here when they build (`docs/scripts/prepare.mjs`),
and the demo has a copy of the favicon.

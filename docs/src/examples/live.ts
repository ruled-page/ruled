/**
 * Brings a page's live examples to life: each `.wb-live` frame gets the
 * editor once it scrolls near the screen, opening its example's workbook
 * with its options, in the docs' colours for the page's light or dark mode.
 */
import type { WorkbookEditorElement } from '@sitepen/spreadsheet';
import { EXAMPLES } from './registry';
import { ledgerTheme, PLAYGROUND_THEMES } from './themes';

let editor: Promise<unknown> | undefined;
/** The standalone build, which defines <workbook-editor>; loaded once, when first needed. */
const loadEditor = () => (editor ??= import(/* @vite-ignore */ `${import.meta.env.BASE_URL.replace(/\/?$/, '/')}editor/workbook-editor.js`));

function start(root: HTMLElement): void {
  const example = EXAMPLES[root.dataset.example ?? ''];
  const host = root.querySelector<HTMLElement>('.wb-live__host');
  const status = root.querySelector<HTMLElement>('.wb-live__status');
  const file = root.querySelector<HTMLElement>('.wb-live__file');
  const received = root.querySelector<HTMLElement>('.wb-live__received');
  if (!example || !host || !status) return;
  const say = (state: string, text: string) => {
    root.dataset.state = state;
    status.textContent = text;
  };
  let theme = ledgerTheme;

  loadEditor()
    .then(() => {
      const el = document.createElement('workbook-editor') as WorkbookEditorElement;
      el.options = {
        warnOnUnload: false,
        drafts: false,
        theme: theme(),
        ...example.options?.((text) => {
          if (received) received.textContent = text;
        }),
      };
      el.addEventListener('opened', (e) => {
        say('ready', 'Live');
        if (file) file.textContent = e.detail.fileName;
      });
      el.addEventListener('editorerror', (e) => say('failed', e.detail.message));
      el.setAttribute('src', sample(example.file));
      host.replaceChildren(el);

      // Follow the page between light and dark.
      new MutationObserver(() => el.setTheme(theme())).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

      for (const button of root.querySelectorAll<HTMLButtonElement>('.wb-live__picker button')) {
        button.addEventListener('click', async () => {
          for (const b of root.querySelectorAll('.wb-live__picker button')) b.setAttribute('aria-pressed', String(b === button));
          if (button.dataset.file) {
            const bytes = await fetch(sample(button.dataset.file)).then((r) => r.arrayBuffer());
            await el.open(bytes, button.dataset.file).catch(() => undefined);
          } else if (button.dataset.theme) {
            theme = PLAYGROUND_THEMES.find((t) => t.id === button.dataset.theme)?.theme ?? ledgerTheme;
            el.setTheme(theme());
          }
        });
      }
    })
    .catch(() => say('failed', "The editor couldn't load"));
}

const sample = (name: string) => `${import.meta.env.BASE_URL.replace(/\/?$/, '/')}samples/${name}`;

/** Starts each live example on the page as it comes near the screen. */
export function startLiveExamples(): void {
  const seen = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        seen.unobserve(entry.target);
        start(entry.target as HTMLElement);
      }
    },
    { rootMargin: '300px' },
  );
  for (const root of document.querySelectorAll<HTMLElement>('.wb-live')) seen.observe(root);
}

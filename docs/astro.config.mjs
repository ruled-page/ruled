// @ts-check
import { fileURLToPath } from 'node:url';
import starlight from '@astrojs/starlight';
import { defineConfig } from 'astro/config';
import starlightLinksValidator from 'starlight-links-validator';
import starlightTypeDoc, { typeDocSidebarGroup } from 'starlight-typedoc';

export default defineConfig({
  // The Cloudflare Pages address; set DOCS_URL for a custom domain.
  site: process.env.DOCS_URL ?? 'https://workbook-editor-docs.pages.dev',
  integrations: [
    starlight({
      title: 'Workbook Editor',
      description: 'Edit real Excel workbooks in the browser, check every cell against a schema with your own async rules, and save .xlsx back intact.',
      logo: { light: './src/assets/logo-light.svg', dark: './src/assets/logo-dark.svg' },
      favicon: '/favicon.svg',
      customCss: ['@fontsource-variable/fraunces', '@fontsource/ibm-plex-sans/400.css', '@fontsource/ibm-plex-sans/500.css', '@fontsource/ibm-plex-sans/600.css', '@fontsource/ibm-plex-mono/400.css', '@fontsource/ibm-plex-mono/500.css', './src/styles/brand.css'],
      components: { Hero: './src/components/landing/Hero.astro' },
      lastUpdated: false,
      pagination: true,
      head: [{ tag: 'link', attrs: { rel: 'alternate', type: 'text/plain', title: 'LLM text', href: '/llms.txt' } }],
      plugins: [
        starlightLinksValidator({ exclude: ['/llms.txt', '/llms-full.txt', '/agents/**', '/**/*.md'] }),
        starlightTypeDoc({
          // Wrappers naming the modules after their import paths.
          entryPoints: ['./api/spreadsheet.ts', './api/standalone.ts'],
          tsconfig: './tsconfig.typedoc.json',
          output: 'reference/api',
          sidebar: { label: 'API reference', collapsed: true },
          typeDoc: {
            skipErrorChecking: true,
            // The source is private, so no "Defined in" links to it.
            disableSources: true,
            // Members inherited from the DOM (the element's HTMLElement ones) are the browser's to document.
            excludeExternals: true,
            readme: './api/README.md',
            excludePrivate: true,
            excludeInternal: true,
            sort: ['kind', 'alphabetical'],
            useCodeBlocks: true,
            parametersFormat: 'table',
            interfacePropertiesFormat: 'table',
          },
        }),
      ],
      sidebar: [
        { label: 'Start here', items: [{ autogenerate: { directory: 'start' } }] },
        { label: 'Schemas and validation', items: [{ autogenerate: { directory: 'schemas' } }] },
        { label: 'Guides', items: [{ autogenerate: { directory: 'guides' } }] },
        { label: 'Live examples', items: [{ autogenerate: { directory: 'examples' } }] },
        { label: 'AI and agents', items: [{ autogenerate: { directory: 'ai' } }] },
        typeDocSidebarGroup,
      ],
    }),
  ],
  vite: {
    resolve: {
      // Examples import the library as hosts do; in the docs that's its types and defineSchema,
      // and the editor itself is the standalone build, loaded when a live example scrolls into view.
      alias: { '@sitepen/spreadsheet': fileURLToPath(new URL('./src/examples/sdk.ts', import.meta.url)) },
    },
  },
});

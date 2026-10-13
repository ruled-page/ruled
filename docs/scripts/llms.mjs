/**
 * Turns the built site into text for language models, after `astro build`:
 *
 *   dist/<page>.md     each page as Markdown (dist/schemas/async-hooks.md)
 *   dist/llms.txt      an index of the pages, in the sidebar's order, as llmstxt.org describes
 *   dist/llms-full.txt every page in one file
 *
 * Pages are read from the built HTML, so what's generated (the API reference)
 * and what's rendered from components (code samples) is included as shown.
 */
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { select, selectAll } from 'hast-util-select';
import rehypeParse from 'rehype-parse';
import rehypeRemark from 'rehype-remark';
import remarkGfm from 'remark-gfm';
import remarkStringify from 'remark-stringify';
import { unified } from 'unified';
import { remove } from 'unist-util-remove';

const dist = fileURLToPath(new URL('../dist/', import.meta.url));
const site = (process.env.DOCS_URL ?? 'https://ruled.page').replace(/\/$/, '');
const siteDescription = 'Edit real Excel workbooks in the browser, check every cell against a schema with your own async rules, and save .xlsx back intact.';

/** Every built page: its route ('schemas/async-hooks/') and its HTML file. */
function pages(dir = dist) {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return name === 'agents' || name === '_astro' ? [] : pages(path);
    if (name !== 'index.html') return [];
    const route = relative(dist, dir).split('\\').join('/');
    return route === '404' ? [] : [{ route: route ? `${route}/` : '', file: path }];
  });
}

const text = (node) => (node.type === 'text' ? node.value : (node.children ?? []).map(text).join(''));

/** A page's title, summary and content as Markdown. */
function convert(html) {
  const tree = unified().use(rehypeParse).parse(html);
  const title = text(select('h1', tree) ?? select('title', tree) ?? { type: 'text', value: '' }).trim();
  // Pages without their own description get the site's, which says nothing about them.
  const own = select('meta[name="description"]', tree)?.properties.content ?? '';
  const description = own === siteDescription ? '' : own;
  const main = select('.sl-markdown-content', tree);
  if (!main) return { title, description, markdown: '' };

  // Code blocks as plain <pre><code>, one line per line, with their language.
  for (const pre of selectAll('pre[data-language]', main)) {
    const lines = selectAll('.ec-line', pre);
    const code = lines.length ? lines.map((l) => text(l)).join('\n') : text(pre);
    pre.children = [{ type: 'element', tagName: 'code', properties: { className: [`language-${pre.properties.dataLanguage}`] }, children: [{ type: 'text', value: code }] }];
  }
  // A live editor becomes a line saying which example it was.
  for (const live of selectAll('.wb-live', main)) {
    const hint = select('.wb-live__hint', live);
    live.tagName = 'p';
    live.children = [{ type: 'element', tagName: 'em', properties: {}, children: [{ type: 'text', value: `Live editor: the ${live.properties.dataExample} example${hint ? `. ${text(hint).trim()}` : '.'}` }] }];
  }
  remove(main, (node) => {
    if (node.type !== 'element') return false;
    const cls = node.properties?.className ?? [];
    // TypeDoc's anchors are empty links.
    const emptyLink = node.tagName === 'a' && !text(node).trim();
    return emptyLink || node.tagName === 'script' || node.tagName === 'style' || node.tagName === 'button' || cls.includes('sl-anchor-link') || cls.includes('copy') || node.tagName === 'figcaption';
  });
  const markdown = unified()
    .use(rehypeRemark)
    // Tables unpadded: lining columns up makes wide ones many times their size.
    .use(remarkGfm, { tablePipeAlign: false, tableCellPadding: false })
    .use(remarkStringify, { bullet: '-', fences: true })
    .stringify(unified().use(rehypeRemark).runSync({ type: 'root', children: [main] }));
  // Links within the site point at the pages' Markdown.
  const linked = markdown.replace(/\]\((\/[^)#\s]*?)\/?(#[^)\s]*)?\)/g, (m, path, hash = '') => (path.match(/\.\w+$/) ? m : `](${site}${path === '' ? '/index' : path}.md${hash})`));
  return { title, description, markdown: linked.trim() };
}

/** The sidebar's groups and links, in order, from a built page that has it. */
function sidebar() {
  const tree = unified().use(rehypeParse).parse(readFileSync(join(dist, 'start/introduction/index.html'), 'utf8'));
  const nav = select('#starlight__sidebar', tree) ?? tree;
  const groups = [];
  for (const details of selectAll('details', nav)) {
    const label = text(select('summary', details) ?? { type: 'text', value: '' }).trim();
    // Nested groups (the API reference's) belong to the outermost.
    if (groups.some((g) => g.node !== details && selectAll('details', g.node).includes(details))) continue;
    groups.push({ label, node: details, links: selectAll('a[href]', details).map((a) => a.properties.href) });
  }
  return groups;
}

const all = new Map();
for (const page of pages()) {
  const converted = convert(readFileSync(page.file, 'utf8'));
  all.set(`/${page.route}`, converted);
  writeFileSync(join(dist, page.route ? `${page.route.replace(/\/$/, '')}.md` : 'index.md'), `# ${converted.title}\n\n${converted.description ? `> ${converted.description}\n\n` : ''}${converted.markdown}\n`);
}

const groups = sidebar();
const listed = new Set();
let index = `# Ruled\n\n> ${siteDescription}\n\n`;
index += 'Ruled (`@ruled/sheet`) is a spreadsheet editor for web apps: a React component and a `<workbook-editor>` element. Schemas describe the workbooks an app expects, with sync and async hooks for checks and business logic. Each link is a page as Markdown; llms-full.txt has them all.\n';
let full = `# Ruled\n\n> ${siteDescription}\n`;
for (const group of groups) {
  const api = /api/i.test(group.label);
  index += `\n## ${api ? 'Optional' : group.label}\n\n`;
  if (api) index += `- [API reference](${site}/reference/api/readme.md): every export of @ruled/sheet, its types and their members\n`;
  for (const href of group.links) {
    const page = all.get(href);
    if (!page || listed.has(href)) continue;
    listed.add(href);
    if (!api) index += `- [${page.title}](${site}${href.replace(/\/$/, '')}.md)${page.description ? `: ${page.description}` : ''}\n`;
    full += `\n---\n\n# ${page.title}\n\nSource: ${site}${href}\n\n${page.markdown}\n`;
  }
}
writeFileSync(join(dist, 'llms.txt'), index);
writeFileSync(join(dist, 'llms-full.txt'), full);
console.log(`Wrote ${all.size} Markdown pages, llms.txt and llms-full.txt (${Math.round(full.length / 1024)} kB).`);

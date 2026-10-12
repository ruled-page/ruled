/**
 * The grid library renders markdown cells with marked, as HTML. The editor
 * never shows markdown cells, so in its place this writes text as text,
 * which keeps the parser, and any HTML a cell could carry, out of the build.
 */
export function marked(text: string): string {
  return text.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
}

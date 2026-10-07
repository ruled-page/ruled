import { parse } from '@babel/parser';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

/**
 * Text people see must come from the catalogue. This parses the editor's
 * screens and the code behind them and lists words written straight in: JSX
 * text, the attributes tooltips and screen readers read, and the text of
 * messages, findings, reasons and confirmations. (Errors for developers,
 * thrown as plain Error, are English by design.)
 */
const FILES = ['src/ui', 'src/editor', 'src/state', 'src/schema', 'src/core', 'src/worker'].flatMap((d) =>
  readdirSync(d)
    .filter((f) => /\.tsx?$/.test(f) && !f.endsWith('.d.ts'))
    .map((f) => join(d, f)),
);

/** Attributes people read or hear. */
const READ = new Set(['aria-label', 'aria-description', 'aria-roledescription', 'aria-valuetext', 'title', 'placeholder', 'alt']);
/** Properties that are shown: confirmation dialogs and the like. */
const SHOWN = new Set(['title', 'body', 'confirmLabel', 'cancelLabel', 'alternateLabel', 'label', 'hint', 'text', 'detail', 'message', 'reason', 'reject', 'rejected']);

/** Not words: symbols, file extensions, CSS and the like. */
const ALLOWED = new Set(['fx', '.xlsx', '.xlsm', '.json']);

type AnyNode = { type: string; [key: string]: unknown };

const words = (s: string) => /[A-Za-z]{2,}/.test(s) && !ALLOWED.has(s.trim());

function literal(node: AnyNode | undefined | null): string | undefined {
  if (!node) return undefined;
  if (node.type === 'StringLiteral') return node.value as string;
  if (node.type === 'TemplateLiteral') return (node.quasis as { value: { cooked: string } }[]).map((q) => q.value.cooked).join('…');
  if (node.type === 'JSXExpressionContainer') return literal(node.expression as AnyNode);
  if (node.type === 'ConditionalExpression') return literal(node.consequent as AnyNode) ?? literal(node.alternate as AnyNode);
  if (node.type === 'LogicalExpression') return literal(node.right as AnyNode);
  return undefined;
}

function hardCoded(source: string, file: string): string[] {
  const found: string[] = [];
  const ast = parse(source, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
  const visit = (node: unknown, parent?: AnyNode): void => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) return node.forEach((n) => visit(n, parent));
    const n = node as AnyNode;
    if (typeof n.type !== 'string') return;
    const line = (n.loc as { start: { line: number } } | undefined)?.start.line;
    const report = (s: string) => found.push(`${file}:${line}: ${s.replace(/\s+/g, ' ').trim()}`);
    if (n.type === 'JSXText' && words(n.value as string)) report(n.value as string);
    if (n.type === 'JSXExpressionContainer' && parent?.type === 'JSXElement') {
      const s = literal(n);
      if (s && words(s)) report(s);
    }
    if (n.type === 'JSXAttribute') {
      const name = (n.name as AnyNode).name as string;
      const s = literal(n.value as AnyNode);
      if (READ.has(name) && s && words(s)) report(`${name}="${s}"`);
    }
    // store.toast('warn', 'text') and the like.
    if (n.type === 'CallExpression') {
      const callee = n.callee as AnyNode;
      const name = callee.type === 'MemberExpression' ? ((callee.property as AnyNode).name as string) : (callee.name as string);
      const args = n.arguments as AnyNode[];
      if (name === 'toast') {
        const s = literal(args[1]);
        if (s && words(s)) report(`toast: ${s}`);
      }
    }
    // { title: 'Delete rows?', confirmLabel: 'Delete' } and the like.
    if (n.type === 'ObjectProperty' && !n.computed) {
      const key = n.key as AnyNode;
      const name = (key.name ?? key.value) as string;
      const s = literal(n.value as AnyNode);
      if (SHOWN.has(name) && s && words(s) && /\s/.test(s)) report(`${name}: ${s}`);
    }
    for (const [k, v] of Object.entries(n)) {
      if (k === 'loc' || k === 'start' || k === 'end' || k === 'extra') continue;
      visit(v, n);
    }
  };
  visit(ast.program);
  return found;
}

describe('translation coverage', () => {
  it('leaves no words written straight into the screens', () => {
    const left = FILES.flatMap((f) => hardCoded(readFileSync(f, 'utf8'), f));
    expect(left).toEqual([]);
  });
});

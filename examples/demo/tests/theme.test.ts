import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { copyTheme, THEME_TOKENS, themeProperty, themeStyle } from '@ruled/sheet/internal/editor/theme.ts';
import { DEMO_THEMES } from '../src/themes';
import { DEFAULT_PALETTE } from '@ruled/sheet/internal/ui/gridPalette.ts';

const css = readFileSync(join(__dirname, '..', '..', '..', 'packages', 'spreadsheet', 'src', 'editor', 'editor.css'), 'utf8');
const rootBlock = css.slice(css.indexOf('.wbe {\n'), css.indexOf('font-family: var(--wbe-font);'));
const declared = [...rootBlock.matchAll(/^ {2}(--wbe-[a-z0-9-]+):\s*([^;]+);/gm)].map((m) => [m[1], m[2].trim()] as const);
const tokens: readonly string[] = THEME_TOKENS;

describe('theme tokens', () => {
  it('are declared in the stylesheet and typed for hosts, one for one', () => {
    expect(tokens.map(themeProperty).sort()).toEqual(declared.map(([name]) => name).sort());
  });

  it('are the only colours in the stylesheet', () => {
    const outside = css.replace(rootBlock, '');
    expect(outside.match(/#[0-9a-fA-F]{3,8}\b|rgba?\(/g) ?? []).toEqual([]);
  });

  it('give the grid the same defaults as the stylesheet', () => {
    const value = (name: string) => declared.find(([n]) => n === name)?.[1];
    for (const [key, v] of Object.entries(DEFAULT_PALETTE)) {
      if (key === 'onSelection') continue;
      expect(value(themeProperty(`grid${key[0].toUpperCase()}${key.slice(1)}`)), key).toBe(v);
    }
  });

  it('turn into inline custom properties', () => {
    expect(themeStyle({ accent: '#c00', gridHeaderBg: '#eee' })).toEqual({ '--wbe-accent': '#c00', '--wbe-grid-header-bg': '#eee' });
    for (const t of DEMO_THEMES) for (const key of Object.keys(t.theme ?? {})) expect(tokens).toContain(key);
  });
});

describe('the popup layer', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('takes every token as the editor resolves it, so cell editors opened outside the editor match it', () => {
    const resolved: Record<string, string> = { '--wbe-accent': '#0f766e', '--wbe-grid-bg': '#1b1f23' };
    vi.stubGlobal('getComputedStyle', () => ({ getPropertyValue: (name: string) => resolved[name] ?? '' }));
    const set = new Map<string, string>();
    const portal = { style: { setProperty: (name: string, value: string) => set.set(name, value) } } as unknown as HTMLElement;
    copyTheme({} as Element, portal);
    expect(set.size).toBe(THEME_TOKENS.length);
    expect(set.get('--wbe-accent')).toBe('#0f766e');
    expect(set.get('--wbe-grid-bg')).toBe('#1b1f23');
    // A token the editor doesn't set is cleared, rather than left from another editor.
    expect(set.get('--wbe-radius')).toBe('');
  });
});

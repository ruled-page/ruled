import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { THEME_TOKENS, themeProperty, themeStyle } from '../src/editor/theme';
import { DEMO_THEMES } from '../src/demo/themes';
import { DEFAULT_PALETTE } from '../src/ui/gridPalette';

const css = readFileSync('src/editor/editor.css', 'utf8');
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

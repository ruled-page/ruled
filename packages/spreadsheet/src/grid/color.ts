/** A colour as rgba(), with its opacity set to `alpha`: from #rgb, #rrggbb, #rrggbbaa, rgb() or rgba(); anything else as it is. */
export function withAlpha(color: string, alpha: number): string {
  const c = color.trim();
  let rgb: number[] | undefined;
  const hex = /^#([0-9a-f]{3,8})$/i.exec(c)?.[1];
  if (hex && (hex.length === 3 || hex.length === 6 || hex.length === 8)) {
    const full = hex.length === 3 ? [...hex].map((h) => h + h).join('') : hex;
    rgb = [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16));
  } else {
    const m = /^rgba?\(([^)]+)\)$/i.exec(c);
    const parts = m?.[1].split(/[\s,/]+/).filter(Boolean);
    if (parts && parts.length >= 3) rgb = parts.slice(0, 3).map(Number);
  }
  return rgb && rgb.every((n) => Number.isFinite(n)) ? `rgba(${rgb.join(', ')}, ${alpha})` : c;
}

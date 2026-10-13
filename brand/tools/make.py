"""
Writes Ruled's marks into brand/ as SVG: the mark, the wordmark and the
lockups, light, dark and single-colour. The letters are Fraunces' own
outlines, made by outline.py, so the files need no fonts to show.

    python3 brand/tools/make.py brand
"""
import json, os, sys

OUT = sys.argv[1]
HERE = os.path.dirname(os.path.abspath(__file__))
w = json.load(open(os.path.join(HERE, 'ruled.json')))
GREEN, INK, PAPER, RED, MINT, NIGHT = '#0B6B4F', '#101A16', '#FBF9F4', '#E0412A', '#3ECF9A', '#0E1411'
RED_DARK = '#FF6B4F'

R = json.load(open(os.path.join(HERE, 'R.json')))
R_CAP = 31.0
R_S = R_CAP / R['capHeight']
R_X = 18 + (42 - R['advance'] * R_S) / 2 + 0.5

def mark(tile, line, margin, ink, uid, size=64):
    """A ledger page: its rules, its red margin line, and an R written on the bottom rule."""
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="{size}" height="{size}" role="img" aria-label="Ruled">
  <defs><clipPath id="{uid}"><rect x="4" y="4" width="56" height="56" rx="14"/></clipPath></defs>
  <rect x="4" y="4" width="56" height="56" rx="14" fill="{tile}"/>
  <g clip-path="url(#{uid})">
    <path d="M4 23H60M4 35H60M4 47H60" stroke="{line}" stroke-width="2.5" opacity=".38"/>
    <path d="M18 4V60" stroke="{margin}" stroke-width="3.5"/>
  </g>
  <path transform="translate({R_X:.2f} 47) scale({R_S:.6f})" fill="{ink}" d="{R['d']}"/>
</svg>
'''

# The wordmark: "Ruled" in Fraunces, on a rule, crossed by the margin line.
CAP = 34.0
s = CAP / w['capHeight']
word_w = w['advance'] * s

def wordmark_group(x, baseline, ink, rule, margin, tick=True):
    rule_y = baseline + 9
    tick_path = f'''
  <path d="M{x - 4.5:.1f} {baseline - CAP - 6:.1f}V{rule_y + 7:.1f}" stroke="{margin}" stroke-width="2.5"/>''' if tick else ''
    return f'''  <path d="M{x - (10 if tick else 2):.1f} {rule_y}H{x + word_w + 6:.1f}" stroke="{rule}" stroke-width="2"/>{tick_path}
  <path transform="translate({x:.2f} {baseline}) scale({s:.6f})" fill="{ink}" d="{w['d']}"/>'''

def wordmark(ink, rule, margin, name):
    width, height = word_w + 26, 64
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width:.0f} {height}" width="{width:.0f}" height="{height}" role="img" aria-label="Ruled">
{wordmark_group(14, 45, ink, rule, margin)}
</svg>
'''

def lockup(tile, line, margin, cell, ink, rule, uid):
    m = mark(tile, line, margin, cell, uid).split('\n', 1)[1].rsplit('</svg>', 1)[0]
    x0 = 64 + 22
    width, height = x0 + word_w + 10, 64
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width:.0f} {height}" width="{width:.0f}" height="{height}" role="img" aria-label="Ruled">
{m}{wordmark_group(x0 - 6, 45, ink, rule, margin, tick=False)}
</svg>
'''

files = {
    'ruled-mark.svg': mark(GREEN, PAPER, RED_DARK, PAPER, 'm'),
    'ruled-mark-dark.svg': mark(MINT, NIGHT, RED, NIGHT, 'm'),
    'ruled-mark-mono.svg': mark(INK, PAPER, PAPER, PAPER, 'm'),
    'ruled-favicon.svg': mark(GREEN, PAPER, RED_DARK, PAPER, 'f'),
    'ruled-wordmark.svg': wordmark(INK, '#C9C3B4', RED, 'w'),
    'ruled-wordmark-dark.svg': wordmark(PAPER, '#3A4A41', RED_DARK, 'w'),
    'ruled-lockup.svg': lockup(GREEN, PAPER, RED_DARK, PAPER, INK, '#C9C3B4', 'l'),
    'ruled-lockup-dark.svg': lockup(MINT, NIGHT, RED, NIGHT, PAPER, '#3A4A41', 'l'),
}
os.makedirs(OUT, exist_ok=True)
for name, svg in files.items():
    open(os.path.join(OUT, name), 'w').write(svg)
print(f'Wrote {len(files)} files to {OUT}')

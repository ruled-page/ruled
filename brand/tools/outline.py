"""
Outlines a word in Fraunces at given axis settings, as one SVG path in font
units (baseline at y=0, y down), rounded to whole units. Needs fontTools and
brotli (pip install fonttools brotli) and the Fraunces variable font, which
the docs install:

    python3 brand/tools/outline.py docs/node_modules/@fontsource-variable/fraunces/files/fraunces-latin-full-normal.woff2 \
      Ruled '{"wght": 640, "opsz": 144, "SOFT": 50, "WONK": 0}' > brand/tools/ruled.json
    python3 brand/tools/outline.py … R '{"wght": 700, "opsz": 144, "SOFT": 50, "WONK": 0}' > brand/tools/R.json
"""
import json, re, sys
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen

src, word = sys.argv[1], sys.argv[2]
axes = json.loads(sys.argv[3])
font = TTFont(src)
if 'fvar' in font:
    print(json.dumps({a.axisTag: [a.minValue, a.defaultValue, a.maxValue] for a in font['fvar'].axes}), file=sys.stderr)
    font = instantiateVariableFont(font, axes)
cmap = font.getBestCmap()
glyphs = font.getGlyphSet()
upm = font['head'].unitsPerEm
x = 0
pen = SVGPathPen(glyphs)
for ch in word:
    name = cmap[ord(ch)]
    glyphs[name].draw(TransformPen(pen, (1, 0, 0, -1, x, 0)))
    x += glyphs[name].width
os2 = font['OS/2']
print(json.dumps({'upm': upm, 'advance': x, 'capHeight': getattr(os2, 'sCapHeight', 0), 'xHeight': getattr(os2, 'sxHeight', 0), 'ascender': font['hhea'].ascent, 'descender': font['hhea'].descent, 'd': re.sub(r'-?\d+\.\d+', lambda m: str(round(float(m.group()))), pen.getCommands())}))

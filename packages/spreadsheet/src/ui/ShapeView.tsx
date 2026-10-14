import { useId, type CSSProperties } from 'react';
import type { ShapeSpec } from '../core/types';

type Point = [number, number];

const poly = (points: Point[]) => `M${points.map(([x, y]) => `${x} ${y}`).join(' L')} Z`;

/**
 * An arrow `len` long and `breadth` across along x, its shaft `shaft` of
 * the breadth thick and its head `head` long, at one end or both; `place`
 * turns it to where it goes.
 */
function arrow(len: number, breadth: number, shaft: number, head: number, both: boolean, place: (p: Point) => Point): string {
  const y1 = (breadth * (1 - shaft)) / 2;
  const y2 = breadth - y1;
  const mid = breadth / 2;
  const at = Math.min(head, both ? len / 2 : len);
  const points: Point[] = both
    ? [[0, mid], [at, 0], [at, y1], [len - at, y1], [len - at, 0], [len, mid], [len - at, breadth], [len - at, y2], [at, y2], [at, breadth]]
    : [[0, y1], [len - at, y1], [len - at, 0], [len, mid], [len - at, breadth], [len - at, y2], [0, y2]];
  return poly(points.map(place));
}

/** The outline of one of Excel's preset shapes in a box `w` by `h`, as an SVG path, with Excel's default adjustments. */
export function shapePath(geometry: string, w: number, h: number, adjust: Record<string, number> = {}): string {
  const ss = Math.min(w, h);
  const adj = (name: string, dflt: number) => (adjust[name] ?? dflt) / 100_000;
  switch (geometry) {
    case 'rect':
    case 'flowChartProcess':
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
    case 'roundRect':
    case 'flowChartAlternateProcess':
    case 'flowChartTerminator': {
      const r = geometry === 'flowChartTerminator' ? ss / 2 : ss * adj('adj', 16_667);
      return `M${r} 0 H${w - r} A${r} ${r} 0 0 1 ${w} ${r} V${h - r} A${r} ${r} 0 0 1 ${w - r} ${h} H${r} A${r} ${r} 0 0 1 0 ${h - r} V${r} A${r} ${r} 0 0 1 ${r} 0 Z`;
    }
    case 'ellipse':
      return `M0 ${h / 2} A${w / 2} ${h / 2} 0 1 1 ${w} ${h / 2} A${w / 2} ${h / 2} 0 1 1 0 ${h / 2} Z`;
    case 'triangle':
      return poly([[w * adj('adj', 50_000), 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond':
    case 'flowChartDecision':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': {
      const o = Math.min(w, ss * adj('adj', 25_000));
      return poly([[o, 0], [w, 0], [w - o, h], [0, h]]);
    }
    case 'trapezoid': {
      const o = Math.min(w / 2, ss * adj('adj', 25_000));
      return poly([[0, h], [o, 0], [w - o, 0], [w, h]]);
    }
    case 'pentagon':
      return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'hexagon': {
      const o = Math.min(w / 2, ss * adj('adj', 25_000));
      return poly([[0, h / 2], [o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h]]);
    }
    case 'octagon': {
      const o = Math.min(w / 2, h / 2, ss * adj('adj', 29_289));
      return poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]);
    }
    case 'homePlate':
    case 'chevron': {
      const o = Math.min(w, ss * adj('adj', 50_000));
      const points: Point[] = [[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]];
      if (geometry === 'chevron') points.push([o, h / 2]);
      return poly(points);
    }
    case 'rightArrow':
      return arrow(w, h, adj('adj1', 50_000), ss * adj('adj2', 50_000), false, ([x, y]) => [x, y]);
    case 'leftArrow':
      return arrow(w, h, adj('adj1', 50_000), ss * adj('adj2', 50_000), false, ([x, y]) => [w - x, y]);
    case 'downArrow':
      return arrow(h, w, adj('adj1', 50_000), ss * adj('adj2', 50_000), false, ([x, y]) => [y, x]);
    case 'upArrow':
      return arrow(h, w, adj('adj1', 50_000), ss * adj('adj2', 50_000), false, ([x, y]) => [y, h - x]);
    case 'leftRightArrow':
      return arrow(w, h, adj('adj1', 50_000), ss * adj('adj2', 50_000), true, ([x, y]) => [x, y]);
    case 'upDownArrow':
      return arrow(h, w, adj('adj1', 50_000), ss * adj('adj2', 50_000), true, ([x, y]) => [y, x]);
    case 'plus': {
      const o = Math.min(w / 2, h / 2, ss * adj('adj', 25_000));
      return poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]);
    }
    case 'star5': {
      const inner = adj('adj', 19_098) / 0.5;
      const points: Point[] = [];
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const r = i % 2 ? inner : 1;
        points.push([w / 2 + (w / 2) * r * Math.cos(a), h / 2 + (h / 2) * r * Math.sin(a)]);
      }
      return poly(points);
    }
    case 'line':
    case 'straightConnector1':
      return `M0 0 L${w} ${h}`;
    default:
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

/** Excel's preset dashes, as multiples of the line's width. */
const DASHES: Record<string, number[]> = {
  dot: [1, 1],
  sysDot: [1, 1],
  dash: [4, 3],
  sysDash: [3, 1],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/**
 * A shape or text box drawn in its box: its outline filled and stroked,
 * turned and flipped, with arrowheads on a line's ends, and its text over
 * it (turned with it, but not flipped, as in Excel).
 */
export default function ShapeView({
  shape,
  style,
  className,
  label,
  font,
  fontStack,
}: {
  shape: ShapeSpec;
  style: CSSProperties;
  className: string;
  label: string;
  /** The workbook's own font, which theme fonts are drawn in. */
  font: string;
  /** A font as CSS, with a stand-in where the device lacks it. */
  fontStack: (family: string) => string;
}) {
  // Marker ids, without the colons React's ids have.
  const id = `wbe-shape${useId().replace(/[^\w-]/g, '')}`;
  const w = Math.max(1, Number(style.width));
  const h = Math.max(1, Number(style.height));
  const line = shape.line;
  const open = shape.geometry === 'line' || shape.geometry === 'straightConnector1';
  const flip = shape.flipH || shape.flipV ? `translate(${shape.flipH ? w : 0} ${shape.flipV ? h : 0}) scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1})` : undefined;
  const marker = (end: 'head' | 'tail', kind: 'arrow' | 'oval') => (
    <marker id={`${id}${end}`} viewBox="0 0 10 10" refX="5" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse">
      {kind === 'oval' ? <circle cx="5" cy="5" r="5" fill={line!.color} /> : <path d="M0 0 L10 5 L0 10 Z" fill={line!.color} />}
    </marker>
  );
  const text = shape.text;
  const [l, t, r, b] = text?.insets ?? [0, 0, 0, 0];
  return (
    <div className={className} style={{ ...style, transform: shape.rotate ? `rotate(${shape.rotate}deg)` : undefined }} role="img" aria-label={label}>
      <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} aria-hidden="true">
        {line?.head || line?.tail ? (
          <defs>
            {line.head && marker('head', line.head)}
            {line.tail && marker('tail', line.tail)}
          </defs>
        ) : null}
        <path
          d={shapePath(shape.geometry, w, h, shape.adjust)}
          transform={flip}
          fill={open ? 'none' : (shape.fill ?? 'none')}
          stroke={line?.color ?? 'none'}
          strokeWidth={line?.width}
          strokeDasharray={line?.dash && DASHES[line.dash] ? DASHES[line.dash].map((n) => n * line.width).join(' ') : undefined}
          strokeLinejoin="round"
          markerStart={line?.head ? `url(#${id}head)` : undefined}
          markerEnd={line?.tail ? `url(#${id}tail)` : undefined}
        />
      </svg>
      {text && (
        <div
          className="wbe-shape-text"
          style={{
            padding: `${t}px ${r}px ${b}px ${l}px`,
            justifyContent: text.anchor === 'middle' ? 'center' : text.anchor === 'bottom' ? 'flex-end' : 'flex-start',
            whiteSpace: text.nowrap ? 'pre' : 'pre-wrap',
            writingMode: text.vertical ? 'vertical-rl' : undefined,
            fontFamily: fontStack(font),
          }}
          aria-hidden="true"
        >
          {text.paragraphs.map((p, i) => (
            <p key={i} style={{ textAlign: p.align, fontSize: p.size }}>
              {p.runs.length
                ? p.runs.map((run, j) =>
                    run.text === '\n' ? (
                      <br key={j} />
                    ) : (
                      <span
                        key={j}
                        style={{
                          fontSize: run.size,
                          fontWeight: run.bold ? 700 : undefined,
                          fontStyle: run.italic ? 'italic' : undefined,
                          textDecoration: run.underline ? 'underline' : undefined,
                          color: run.color,
                          fontFamily: run.font && fontStack(run.font),
                        }}
                      >
                        {run.text}
                      </span>
                    ),
                  )
                : ' '}
            </p>
          ))}
        </div>
      )}
    </div>
  );
}

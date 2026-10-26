// SPDX-License-Identifier: MIT
// Shape outlines: our own implementations of the common DrawingML preset shapes (with their adjust handles),
// custom geometry with its guide formulas, and conversion to SVG path data. Everything is in CSS pixels
// relative to the shape box.
import { attr, child, children, clamp, int } from './xml.js';

export type Cmd =
  | { t: 'M'; x: number; y: number }
  | { t: 'L'; x: number; y: number }
  | { t: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { t: 'Z' };

export interface ShapePath {
  cmds: Cmd[];
  /** norm, none, or one of the shading modes (darken, lighten …). */
  fill: string;
  stroke: boolean;
}

export interface Rect {
  l: number;
  t: number;
  r: number;
  b: number;
}

export interface Geometry {
  paths: ShapePath[];
  /** Text rectangle inside the box; null means the whole box. */
  text: Rect | null;
  /** Lines and open connectors: arrowheads apply, nothing to fill. */
  open?: boolean;
}

const DEG = Math.PI / 180;
// Safety net against hostile custom geometry.
const MAX_COMMANDS = 20000;

class PathBuilder {
  readonly paths: ShapePath[] = [];
  private cmds: Cmd[] = [];
  private x = 0;
  private y = 0;
  private count = 0;

  begin(fill = 'norm', stroke = true): this {
    this.cmds = [];
    this.paths.push({ cmds: this.cmds, fill, stroke });
    return this;
  }

  private push(cmd: Cmd): void {
    if (++this.count > MAX_COMMANDS) return;
    if (!this.paths.length) this.begin();
    this.cmds.push(cmd);
  }

  M(x: number, y: number): this {
    this.push({ t: 'M', x, y });
    this.x = x;
    this.y = y;
    return this;
  }

  L(x: number, y: number): this {
    this.push({ t: 'L', x, y });
    this.x = x;
    this.y = y;
    return this;
  }

  C(x1: number, y1: number, x2: number, y2: number, x: number, y: number): this {
    this.push({ t: 'C', x1, y1, x2, y2, x, y });
    this.x = x;
    this.y = y;
    return this;
  }

  Q(x1: number, y1: number, x: number, y: number): this {
    const { x: x0, y: y0 } = this;
    return this.C(x0 + ((x1 - x0) * 2) / 3, y0 + ((y1 - y0) * 2) / 3, x + ((x1 - x) * 2) / 3, y + ((y1 - y) * 2) / 3, x, y);
  }

  Z(): this {
    this.push({ t: 'Z' });
    return this;
  }

  /**
   * DrawingML arcTo: the current point lies on an ellipse with radii wR/hR at the visual angle stAng,
   * and the arc sweeps swAng (degrees, clockwise because y points down).
   */
  arc(wR: number, hR: number, stAng: number, swAng: number): this {
    if (!swAng) return this;
    if (wR <= 0 || hR <= 0) {
      // Degenerate ellipse: the arc collapses onto a line to its end point.
      const end = (stAng + swAng) * DEG;
      const st = stAng * DEG;
      return this.L(this.x + wR * (Math.cos(end) - Math.cos(st)), this.y + hR * (Math.sin(end) - Math.sin(st)));
    }
    const param = (deg: number) => Math.atan2(wR * Math.sin(deg * DEG), hR * Math.cos(deg * DEG));
    const t0 = param(stAng);
    const cx = this.x - wR * Math.cos(t0);
    const cy = this.y - hR * Math.sin(t0);
    let dt = param(stAng + swAng) - t0;
    if (Math.abs(swAng) >= 360) dt = Math.sign(swAng) * 2 * Math.PI;
    else if (swAng > 0) while (dt <= 1e-9) dt += 2 * Math.PI;
    else while (dt >= -1e-9) dt -= 2 * Math.PI;
    const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
    const step = dt / n;
    const k = (4 / 3) * Math.tan(step / 4);
    for (let i = 0; i < n; i++) {
      const a1 = t0 + step * i;
      const a2 = a1 + step;
      const x1 = cx + wR * Math.cos(a1), y1 = cy + hR * Math.sin(a1);
      const x2 = cx + wR * Math.cos(a2), y2 = cy + hR * Math.sin(a2);
      this.C(x1 - k * wR * Math.sin(a1), y1 + k * hR * Math.cos(a1), x2 + k * wR * Math.sin(a2), y2 - k * hR * Math.cos(a2), x2, y2);
    }
    return this;
  }

  ellipse(cx: number, cy: number, rx: number, ry: number): this {
    return this.M(cx - rx, cy).arc(rx, ry, 180, 360).Z();
  }

  poly(points: [number, number][], close = true): this {
    points.forEach(([x, y], i) => (i ? this.L(x, y) : this.M(x, y)));
    return close ? this.Z() : this;
  }
}

type Adj = (name: string, fallback: number) => number;
type Preset = (p: PathBuilder, w: number, h: number, adj: Adj) => Rect | null | void;

const ss = (w: number, h: number) => Math.min(w, h);

function roundRect(p: PathBuilder, w: number, h: number, r: number): void {
  p.M(0, r).arc(r, r, 180, 90).L(w - r, 0).arc(r, r, 270, 90).L(w, h - r).arc(r, r, 0, 90).L(r, h).arc(r, r, 90, 90).Z();
}

/** Points of a regular polygon (or star when inner > 0), stretched so the outer points fill the box. */
function regular(n: number, w: number, h: number, start: number, inner = 0): [number, number][] {
  const raw: [number, number][] = [];
  for (let i = 0; i < n; i++) {
    const a = (start + (360 / n) * i) * DEG;
    raw.push([Math.cos(a), Math.sin(a)]);
    if (inner > 0) {
      const b = (start + (360 / n) * (i + 0.5)) * DEG;
      raw.push([Math.cos(b) * inner, Math.sin(b) * inner]);
    }
  }
  const xs = raw.map(([x]) => x), ys = raw.map(([, y]) => y);
  const minX = Math.min(...xs), maxX = Math.max(...xs), minY = Math.min(...ys), maxY = Math.max(...ys);
  return raw.map(([x, y]) => [((x - minX) / (maxX - minX || 1)) * w, ((y - minY) / (maxY - minY || 1)) * h]);
}

function arrowRight(p: PathBuilder, w: number, h: number, adj: Adj): Rect {
  const a1 = clamp(adj('adj1', 50000), 0, 100000);
  const a2 = clamp(adj('adj2', 50000), 0, (100000 * w) / ss(w, h));
  const dy1 = (h * a1) / 200000;
  const dx1 = (ss(w, h) * a2) / 100000;
  const x1 = w - dx1, y1 = h / 2 - dy1, y2 = h / 2 + dy1;
  p.poly([[0, y1], [x1, y1], [x1, 0], [w, h / 2], [x1, h], [x1, y2], [0, y2]]);
  return { l: 0, t: y1, r: x1 + (y1 * dx1) / (h / 2 || 1), b: y2 };
}

function arrowUp(p: PathBuilder, w: number, h: number, adj: Adj): Rect {
  const a1 = clamp(adj('adj1', 50000), 0, 100000);
  const a2 = clamp(adj('adj2', 50000), 0, (100000 * h) / ss(w, h));
  const dx1 = (w * a1) / 200000;
  const dy1 = (ss(w, h) * a2) / 100000;
  const x1 = w / 2 - dx1, x2 = w / 2 + dx1;
  p.poly([[0, dy1], [w / 2, 0], [w, dy1], [x2, dy1], [x2, h], [x1, h], [x1, dy1]]);
  return { l: x1, t: dy1 - (x1 * dy1) / (w / 2 || 1), r: x2, b: h };
}

/** Mirrors whatever a preset drew horizontally and/or vertically (left arrow from right arrow …). */
function mirrored(base: Preset, flipX: boolean, flipY: boolean): Preset {
  return (p, w, h, adj) => {
    const start = p.paths.length;
    const rect = base(p, w, h, adj) ?? null;
    const fx = (x: number) => (flipX ? w - x : x), fy = (y: number) => (flipY ? h - y : y);
    for (const path of p.paths.slice(start)) {
      for (const c of path.cmds) {
        if (c.t === 'Z') continue;
        c.x = fx(c.x);
        c.y = fy(c.y);
        if (c.t === 'C') {
          c.x1 = fx(c.x1);
          c.y1 = fy(c.y1);
          c.x2 = fx(c.x2);
          c.y2 = fy(c.y2);
        }
      }
    }
    if (!rect) return null;
    const [l, r] = flipX ? [w - rect.r, w - rect.l] : [rect.l, rect.r];
    const [t, b] = flipY ? [h - rect.b, h - rect.t] : [rect.t, rect.b];
    return { l, t, r, b };
  };
}

function wedgeCallout(p: PathBuilder, w: number, h: number, adj: Adj, corner: number): void {
  const xPos = w / 2 + (w * adj('adj1', -20833)) / 100000;
  const yPos = h / 2 + (h * adj('adj2', 62500)) / 100000;
  const r = corner;
  const inside = xPos >= 0 && xPos <= w && yPos >= 0 && yPos <= h;
  const dx = xPos - w / 2, dy = yPos - h / 2;
  const vertical = Math.abs(dy) * w > Math.abs(dx) * h;
  // The wedge leaves the side facing the tip, from the third of that side nearer to the tip.
  const along = (len: number, toward: number) => (toward > len / 2 ? [(len * 7) / 12, (len * 10) / 12] : [(len * 2) / 12, (len * 5) / 12]);
  p.M(0, r);
  if (r) p.arc(r, r, 180, 90);
  if (!inside && vertical && dy < 0) {
    const [a, b] = along(w, xPos);
    p.L(a, 0).L(xPos, yPos).L(b, 0);
  }
  p.L(w - r, 0);
  if (r) p.arc(r, r, 270, 90);
  if (!inside && !vertical && dx > 0) {
    const [a, b] = along(h, yPos);
    p.L(w, a).L(xPos, yPos).L(w, b);
  }
  p.L(w, h - r);
  if (r) p.arc(r, r, 0, 90);
  if (!inside && vertical && dy > 0) {
    const [a, b] = along(w, xPos);
    p.L(b, h).L(xPos, yPos).L(a, h);
  }
  p.L(r, h);
  if (r) p.arc(r, r, 90, 90);
  if (!inside && !vertical && dx < 0) {
    const [a, b] = along(h, yPos);
    p.L(0, b).L(xPos, yPos).L(0, a);
  }
  p.Z();
}

function callout(defaults: number[], border: boolean, accent: boolean): Preset {
  return (p, w, h, adj) => {
    p.begin('norm', border).poly([[0, 0], [w, 0], [w, h], [0, h]]);
    const pts: [number, number][] = [];
    for (let i = 0; i < defaults.length; i += 2) {
      pts.push([(w * adj(`adj${i + 2}`, defaults[i + 1])) / 100000, (h * adj(`adj${i + 1}`, defaults[i])) / 100000]);
    }
    if (accent) p.begin('none', true).poly([[pts[0][0], 0], [pts[0][0], h]], false);
    p.begin('none', true).poly(pts, false);
  };
}

function star(n: number, fallback: number, start = -90): Preset {
  return (p, w, h, adj) => {
    const inner = clamp(adj('adj', fallback), 0, 50000) / 50000;
    p.poly(regular(n, w, h, start, inner));
    return { l: w * (0.5 - inner * 0.35), t: h * (0.5 - inner * 0.35), r: w * (0.5 + inner * 0.35), b: h * (0.5 + inner * 0.35) };
  };
}

const ELLIPSE_INSET = (1 - Math.SQRT1_2) / 2;

/** "[" opening to the right; the right bracket is its mirror image. */
function bracket(p: PathBuilder, w: number, h: number, adj: Adj): void {
  const y1 = Math.min((ss(w, h) * adj('adj', 8333)) / 100000, h / 2);
  p.begin('norm', false).M(w, h).arc(w, y1, 90, 90).L(0, y1).arc(w, y1, 180, 90).Z();
  p.begin('none').M(w, h).arc(w, y1, 90, 90).L(0, y1).arc(w, y1, 180, 90);
}

/** "{" with its point at adj2 of the height; drawn as an outline only. */
function brace(p: PathBuilder, w: number, h: number, adj: Adj): void {
  const y1 = Math.min((ss(w, h) * adj('adj1', 8333)) / 100000, h / 4);
  const ym = (h * clamp(adj('adj2', 50000), 0, 100000)) / 100000;
  p.begin('none').M(w, h).arc(w / 2, y1, 90, 90).L(w / 2, ym + y1).arc(w / 2, y1, 0, -90).arc(w / 2, y1, 90, -90).L(w / 2, y1).arc(w / 2, y1, 180, 90);
}

const PRESETS: Record<string, Preset> = {
  rect: (p, w, h) => void p.poly([[0, 0], [w, 0], [w, h], [0, h]]),
  roundRect: (p, w, h, adj) => {
    const r = (ss(w, h) * clamp(adj('adj', 16667), 0, 50000)) / 100000;
    roundRect(p, w, h, r);
    const i = r * 0.29289;
    return { l: i, t: i, r: w - i, b: h - i };
  },
  ellipse: (p, w, h) => {
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    return { l: w * ELLIPSE_INSET, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  line: (p, w, h) => void p.begin('none').M(0, 0).L(w, h),
  lineInv: (p, w, h) => void p.begin('none').M(0, h).L(w, 0),
  straightConnector1: (p, w, h) => void p.begin('none').M(0, 0).L(w, h),
  bentConnector2: (p, w, h) => void p.begin('none').M(0, 0).L(w, 0).L(w, h),
  bentConnector3: (p, w, h, adj) => {
    const x1 = (w * adj('adj1', 50000)) / 100000;
    p.begin('none').M(0, 0).L(x1, 0).L(x1, h).L(w, h);
  },
  bentConnector4: (p, w, h, adj) => {
    const x1 = (w * adj('adj1', 50000)) / 100000, y2 = (h * adj('adj2', 50000)) / 100000;
    p.begin('none').M(0, 0).L(x1, 0).L(x1, y2).L(w, y2).L(w, h);
  },
  bentConnector5: (p, w, h, adj) => {
    const x1 = (w * adj('adj1', 50000)) / 100000, y2 = (h * adj('adj2', 50000)) / 100000, x3 = (w * adj('adj3', 50000)) / 100000;
    p.begin('none').M(0, 0).L(x1, 0).L(x1, y2).L(x3, y2).L(x3, h).L(w, h);
  },
  curvedConnector2: (p, w, h) => void p.begin('none').M(0, 0).C(w / 2, 0, w, h / 2, w, h),
  curvedConnector3: (p, w, h, adj) => {
    const x2 = (w * adj('adj1', 50000)) / 100000;
    p.begin('none').M(0, 0).C(x2 / 2, 0, x2, h / 4, x2, h / 2).C(x2, (h * 3) / 4, (w + x2) / 2, h, w, h);
  },
  curvedConnector4: (p, w, h, adj) => {
    const x2 = (w * adj('adj1', 50000)) / 100000, y4 = (h * adj('adj2', 50000)) / 100000;
    p.begin('none').M(0, 0).C(x2 / 2, 0, x2, y4 / 2, x2, y4 / 2).C(x2, y4, (x2 + w) / 2, y4, (x2 + w) / 2, y4).C(w, y4, w, (y4 + h) / 2, w, h);
  },
  curvedConnector5: (p, w, h, adj) => {
    const x1 = (w * adj('adj1', 50000)) / 100000, y2 = (h * adj('adj2', 50000)) / 100000, x3 = (w * adj('adj3', 50000)) / 100000;
    p.begin('none').M(0, 0).C(x1 / 2, 0, x1, y2 / 2, x1, y2 / 2).C(x1, y2, x3, y2, x3, (y2 + h) / 2).C(x3, h, w, h, w, h);
  },
  triangle: (p, w, h, adj) => {
    const x1 = (w * clamp(adj('adj', 50000), 0, 100000)) / 100000;
    p.poly([[0, h], [x1, 0], [w, h]]);
    return { l: x1 / 2, t: h / 2, r: x1 / 2 + w / 2, b: h };
  },
  rtTriangle: (p, w, h) => {
    p.poly([[0, h], [0, 0], [w, h]]);
    return { l: w / 12, t: (h * 7) / 12, r: (w * 7) / 12, b: (h * 11) / 12 };
  },
  diamond: (p, w, h) => {
    p.poly([[0, h / 2], [w / 2, 0], [w, h / 2], [w / 2, h]]);
    return { l: w / 4, t: h / 4, r: (w * 3) / 4, b: (h * 3) / 4 };
  },
  parallelogram: (p, w, h, adj) => {
    const x2 = (ss(w, h) * clamp(adj('adj', 25000), 0, (100000 * w) / ss(w, h))) / 100000;
    p.poly([[0, h], [x2, 0], [w, 0], [w - x2, h]]);
    return { l: x2 / 2, t: (h * x2) / (w * 2 || 1), r: w - x2 / 2, b: h - (h * x2) / (w * 2 || 1) };
  },
  trapezoid: (p, w, h, adj) => {
    const x2 = (ss(w, h) * clamp(adj('adj', 25000), 0, (50000 * w) / ss(w, h))) / 100000;
    p.poly([[0, h], [x2, 0], [w - x2, 0], [w, h]]);
    return { l: (x2 * 2) / 3, t: (x2 * 2) / 3 > h ? 0 : (h * x2) / (w * 1.5 || 1), r: w - (x2 * 2) / 3, b: h };
  },
  pentagon: (p, w, h) => {
    const pts = regular(5, w, h, -90);
    p.poly(pts);
    // Between the lower corners, from a bit above the side corners down to the bottom.
    return { l: pts[3][0], t: pts[1][1] * 0.618, r: pts[2][0], b: h };
  },
  hexagon: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 25000), 0, (50000 * w) / ss(w, h))) / 100000;
    p.poly([[0, h / 2], [x1, 0], [w - x1, 0], [w, h / 2], [w - x1, h], [x1, h]]);
    return { l: x1 / 2, t: h / 8, r: w - x1 / 2, b: (h * 7) / 8 };
  },
  heptagon: (p, w, h) => {
    p.poly(regular(7, w, h, -90));
    return { l: w * 0.15, t: h * 0.2, r: w * 0.85, b: h * 0.85 };
  },
  octagon: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 29289), 0, 50000)) / 100000;
    p.poly([[0, x1], [x1, 0], [w - x1, 0], [w, x1], [w, h - x1], [w - x1, h], [x1, h], [0, h - x1]]);
    return { l: x1 / 2, t: x1 / 2, r: w - x1 / 2, b: h - x1 / 2 };
  },
  decagon: (p, w, h) => {
    p.poly(regular(10, w, h, 0));
    return { l: w * 0.1, t: h * 0.15, r: w * 0.9, b: h * 0.85 };
  },
  dodecagon: (p, w, h) => {
    p.poly(regular(12, w, h, 15));
    return { l: w * 0.13, t: h * 0.13, r: w * 0.87, b: h * 0.87 };
  },
  chevron: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 50000), 0, (100000 * w) / ss(w, h))) / 100000;
    p.poly([[0, 0], [w - x1, 0], [w, h / 2], [w - x1, h], [0, h], [x1, h / 2]]);
    return { l: Math.min(x1, w - x1), t: 0, r: Math.max(x1, w - x1), b: h };
  },
  homePlate: (p, w, h, adj) => {
    const x1 = w - (ss(w, h) * clamp(adj('adj', 50000), 0, (100000 * w) / ss(w, h))) / 100000;
    p.poly([[0, 0], [x1, 0], [w, h / 2], [x1, h], [0, h]]);
    return { l: 0, t: 0, r: (x1 + w) / 2, b: h };
  },
  rightArrow: arrowRight,
  leftArrow: mirrored(arrowRight, true, false),
  upArrow: arrowUp,
  downArrow: mirrored(arrowUp, false, true),
  notchedRightArrow: (p, w, h, adj) => {
    const a1 = clamp(adj('adj1', 50000), 0, 100000);
    const dy1 = (h * a1) / 200000, dx1 = (ss(w, h) * clamp(adj('adj2', 50000), 0, (100000 * w) / ss(w, h))) / 100000;
    const x1 = w - dx1, y1 = h / 2 - dy1, y2 = h / 2 + dy1, notch = (dy1 * dx1) / (h / 2 || 1);
    p.poly([[0, y1], [x1, y1], [x1, 0], [w, h / 2], [x1, h], [x1, y2], [0, y2], [notch, h / 2]]);
    return { l: notch, t: y1, r: x1, b: y2 };
  },
  stripedRightArrow: arrowRight,
  leftRightArrow: (p, w, h, adj) => {
    const dy1 = (h * clamp(adj('adj1', 50000), 0, 100000)) / 200000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 50000), 0, (50000 * w) / ss(w, h))) / 100000;
    const x3 = w - x2, y1 = h / 2 - dy1, y2 = h / 2 + dy1;
    p.poly([[0, h / 2], [x2, 0], [x2, y1], [x3, y1], [x3, 0], [w, h / 2], [x3, h], [x3, y2], [x2, y2], [x2, h]]);
    return { l: x2 / 2, t: y1, r: w - x2 / 2, b: y2 };
  },
  upDownArrow: (p, w, h, adj) => {
    const dx1 = (w * clamp(adj('adj1', 50000), 0, 100000)) / 200000;
    const y2 = (ss(w, h) * clamp(adj('adj2', 50000), 0, (50000 * h) / ss(w, h))) / 100000;
    const y3 = h - y2, x1 = w / 2 - dx1, x2 = w / 2 + dx1;
    p.poly([[0, y2], [w / 2, 0], [w, y2], [x2, y2], [x2, y3], [w, y3], [w / 2, h], [0, y3], [x1, y3], [x1, y2]]);
    return { l: x1, t: y2 / 2, r: x2, b: h - y2 / 2 };
  },
  star4: star(4, 12500),
  star5: star(5, 19098),
  star6: star(6, 28868, 0),
  star7: star(7, 34601),
  star8: star(8, 38250),
  star10: star(10, 42533, 0),
  star12: star(12, 37500),
  star16: star(16, 37500),
  star24: star(24, 37500),
  star32: star(32, 37500),
  plus: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 25000), 0, 50000)) / 100000;
    const x2 = w - x1, y2 = h - x1;
    p.poly([[0, x1], [x1, x1], [x1, 0], [x2, 0], [x2, x1], [w, x1], [w, y2], [x2, y2], [x2, h], [x1, h], [x1, y2], [0, y2]]);
    return { l: 0, t: x1, r: w, b: y2 };
  },
  donut: (p, w, h, adj) => {
    const dr = (ss(w, h) * clamp(adj('adj', 25000), 0, 50000)) / 100000;
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    // Reverse direction for the hole so it stays empty under the nonzero fill rule as well.
    p.M(dr, h / 2).arc(w / 2 - dr, h / 2 - dr, 180, -360).Z();
    return { l: w * ELLIPSE_INSET, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  noSmoking: (p, w, h, adj) => {
    const dr = (ss(w, h) * clamp(adj('adj', 18750), 0, 50000)) / 100000;
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    p.M(dr, h / 2).arc(w / 2 - dr, h / 2 - dr, 180, -360).Z();
    const c = Math.SQRT1_2, rx = w / 2 - dr, ry = h / 2 - dr, t = dr / 2;
    p.poly([[w / 2 - rx * c - t, h / 2 - ry * c + t], [w / 2 - rx * c + t, h / 2 - ry * c - t], [w / 2 + rx * c + t, h / 2 + ry * c - t], [w / 2 + rx * c - t, h / 2 + ry * c + t]]);
  },
  frame: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 12500), 0, 50000)) / 100000;
    p.poly([[0, 0], [w, 0], [w, h], [0, h]]);
    p.poly([[x1, x1], [x1, h - x1], [w - x1, h - x1], [w - x1, x1]]);
    return { l: x1, t: x1, r: w - x1, b: h - x1 };
  },
  halfFrame: (p, w, h, adj) => {
    const y1 = (ss(w, h) * clamp(adj('adj1', 33333), 0, 100000)) / 100000;
    const x1 = (ss(w, h) * clamp(adj('adj2', 33333), 0, 100000)) / 100000;
    p.poly([[0, 0], [w, 0], [w - (y1 * w) / (h || 1), y1], [x1, y1], [x1, h - (x1 * h) / (w || 1)], [0, h]]);
  },
  corner: (p, w, h, adj) => {
    const dy1 = (ss(w, h) * clamp(adj('adj1', 50000), 0, 100000)) / 100000;
    const dx1 = (ss(w, h) * clamp(adj('adj2', 50000), 0, 100000)) / 100000;
    p.poly([[0, 0], [dx1, 0], [dx1, h - dy1], [w, h - dy1], [w, h], [0, h]]);
  },
  diagStripe: (p, w, h, adj) => {
    const a = clamp(adj('adj', 50000), 0, 100000) / 100000;
    p.poly([[0, h * a], [w * a, 0], [w, 0], [0, h]]);
  },
  plaque: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 16667), 0, 50000)) / 100000;
    p.M(0, x1).arc(x1, x1, 90, -90).L(w - x1, 0).arc(x1, x1, 180, -90).L(w, h - x1).arc(x1, x1, 270, -90).L(x1, h).arc(x1, x1, 0, -90).Z();
    return { l: x1 * 0.7, t: x1 * 0.7, r: w - x1 * 0.7, b: h - x1 * 0.7 };
  },
  snip1Rect: (p, w, h, adj) => {
    const d = (ss(w, h) * clamp(adj('adj', 16667), 0, 50000)) / 100000;
    p.poly([[0, 0], [w - d, 0], [w, d], [w, h], [0, h]]);
    return { l: 0, t: d / 2, r: w - d / 2, b: h };
  },
  snip2SameRect: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 16667), 0, 50000)) / 100000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 0), 0, 50000)) / 100000;
    p.poly([[x1, 0], [w - x1, 0], [w, x1], [w, h - x2], [w - x2, h], [x2, h], [0, h - x2], [0, x1]]);
    return { l: Math.max(x1, x2) / 2, t: x1 / 2, r: w - Math.max(x1, x2) / 2, b: h - x2 / 2 };
  },
  snip2DiagRect: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 0), 0, 50000)) / 100000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 16667), 0, 50000)) / 100000;
    p.poly([[x1, 0], [w - x2, 0], [w, x2], [w, h - x1], [w - x1, h], [x2, h], [0, h - x2], [0, x1]]);
  },
  snipRoundRect: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 16667), 0, 50000)) / 100000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 16667), 0, 50000)) / 100000;
    p.M(x1, 0).L(w - x2, 0).L(w, x2).L(w, h).L(0, h).L(0, x1).arc(x1, x1, 180, 90).Z();
  },
  round1Rect: (p, w, h, adj) => {
    const d = (ss(w, h) * clamp(adj('adj', 16667), 0, 50000)) / 100000;
    p.M(0, 0).L(w - d, 0).arc(d, d, 270, 90).L(w, h).L(0, h).Z();
    return { l: 0, t: 0, r: w - d * 0.29289, b: h };
  },
  round2SameRect: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 16667), 0, 50000)) / 100000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 0), 0, 50000)) / 100000;
    p.M(x1, 0).L(w - x1, 0).arc(x1, x1, 270, 90).L(w, h - x2).arc(x2, x2, 0, 90).L(x2, h).arc(x2, x2, 90, 90).L(0, x1).arc(x1, x1, 180, 90).Z();
    const i = Math.max(x1, x2) * 0.29289;
    return { l: i, t: x1 * 0.29289, r: w - i, b: h - x2 * 0.29289 };
  },
  round2DiagRect: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj1', 16667), 0, 50000)) / 100000;
    const x2 = (ss(w, h) * clamp(adj('adj2', 0), 0, 50000)) / 100000;
    p.M(x1, 0).L(w - x2, 0).arc(x2, x2, 270, 90).L(w, h - x1).arc(x1, x1, 0, 90).L(x2, h).arc(x2, x2, 90, 90).L(0, x1).arc(x1, x1, 180, 90).Z();
  },
  flowChartProcess: (p, w, h) => void p.poly([[0, 0], [w, 0], [w, h], [0, h]]),
  flowChartAlternateProcess: (p, w, h) => {
    const r = ss(w, h) / 6;
    roundRect(p, w, h, r);
    return { l: r * 0.29289, t: r * 0.29289, r: w - r * 0.29289, b: h - r * 0.29289 };
  },
  flowChartDecision: (p, w, h) => {
    p.poly([[0, h / 2], [w / 2, 0], [w, h / 2], [w / 2, h]]);
    return { l: w / 4, t: h / 4, r: (w * 3) / 4, b: (h * 3) / 4 };
  },
  flowChartTerminator: (p, w, h) => {
    const rx = (w * 3475) / 21600;
    p.M(rx, 0).L(w - rx, 0).arc(rx, h / 2, 270, 180).L(rx, h).arc(rx, h / 2, 90, 180).Z();
    return { l: (w * 1018) / 21600, t: (h * 3163) / 21600, r: (w * 20582) / 21600, b: (h * 18437) / 21600 };
  },
  flowChartDocument: (p, w, h) => {
    const y = (k: number) => (h * k) / 21600;
    p.M(0, 0).L(w, 0).L(w, y(17322)).C(w / 2, y(17322), w / 2, y(23922), 0, y(20172)).Z();
    return { l: 0, t: 0, r: w, b: y(17322) };
  },
  flowChartMultidocument: (p, w, h) => {
    const y = (k: number) => (h * k) / 21600;
    p.M(0, y(3675)).L(w * 0.82, y(3675)).L(w * 0.82, y(18022)).C(w * 0.41, y(18022), w * 0.41, y(23400), 0, y(20400)).Z();
    p.begin('norm').poly([[w * 0.06, y(3675)], [w * 0.06, y(1815)], [w * 0.9, y(1815)], [w * 0.9, y(16252)], [w * 0.82, y(16252)]], false);
    p.begin('norm').poly([[w * 0.12, y(1815)], [w * 0.12, 0], [w, 0], [w, y(14392)], [w * 0.9, y(14392)]], false);
  },
  flowChartInputOutput: (p, w, h) => {
    p.poly([[0, h], [w / 5, 0], [w, 0], [(w * 4) / 5, h]]);
    return { l: w / 5, t: 0, r: (w * 4) / 5, b: h };
  },
  flowChartPredefinedProcess: (p, w, h) => {
    p.poly([[0, 0], [w, 0], [w, h], [0, h]]);
    p.begin('none').M(w / 8, 0).L(w / 8, h).M((w * 7) / 8, 0).L((w * 7) / 8, h);
    return { l: w / 8, t: 0, r: (w * 7) / 8, b: h };
  },
  flowChartInternalStorage: (p, w, h) => {
    p.poly([[0, 0], [w, 0], [w, h], [0, h]]);
    p.begin('none').M(w / 8, 0).L(w / 8, h).M(0, h / 8).L(w, h / 8);
    return { l: w / 8, t: h / 8, r: w, b: h };
  },
  flowChartConnector: (p, w, h) => {
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    return { l: w * ELLIPSE_INSET, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  flowChartManualInput: (p, w, h) => {
    p.poly([[0, h / 5], [w, 0], [w, h], [0, h]]);
    return { l: 0, t: h / 5, r: w, b: h };
  },
  flowChartManualOperation: (p, w, h) => {
    p.poly([[0, 0], [w, 0], [(w * 4) / 5, h], [w / 5, h]]);
    return { l: w / 5, t: 0, r: (w * 4) / 5, b: h };
  },
  flowChartPreparation: (p, w, h) => {
    p.poly([[0, h / 2], [w / 5, 0], [(w * 4) / 5, 0], [w, h / 2], [(w * 4) / 5, h], [w / 5, h]]);
    return { l: w / 5, t: 0, r: (w * 4) / 5, b: h };
  },
  flowChartMerge: (p, w, h) => {
    p.poly([[0, 0], [w, 0], [w / 2, h]]);
    return { l: w / 4, t: 0, r: (w * 3) / 4, b: h / 2 };
  },
  flowChartExtract: (p, w, h) => {
    p.poly([[w / 2, 0], [w, h], [0, h]]);
    return { l: w / 4, t: h / 2, r: (w * 3) / 4, b: h };
  },
  flowChartOffpageConnector: (p, w, h) => {
    p.poly([[0, 0], [w, 0], [w, (h * 4) / 5], [w / 2, h], [0, (h * 4) / 5]]);
    return { l: 0, t: 0, r: w, b: (h * 4) / 5 };
  },
  flowChartPunchedCard: (p, w, h) => {
    p.poly([[0, h / 5], [w / 5, 0], [w, 0], [w, h], [0, h]]);
    return { l: 0, t: h / 5, r: w, b: h };
  },
  flowChartPunchedTape: (p, w, h) => {
    const a = h / 10;
    p.M(0, a).C(w / 4, -a, w / 4, 3 * a, w / 2, a).C((w * 3) / 4, -a, (w * 3) / 4, 3 * a, w, a).L(w, h - a);
    p.C((w * 3) / 4, h - 3 * a, (w * 3) / 4, h + a, w / 2, h - a).C(w / 4, h - 3 * a, w / 4, h + a, 0, h - a).Z();
    return { l: 0, t: h / 5, r: w, b: (h * 4) / 5 };
  },
  flowChartDelay: (p, w, h) => {
    p.M(0, 0).L(w / 2, 0).arc(w / 2, h / 2, 270, 180).L(0, h).Z();
    return { l: 0, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  flowChartDisplay: (p, w, h) => {
    p.M(0, h / 2).L(w / 6, 0).L((w * 5) / 6, 0).arc(w / 6, h / 2, 270, 180).L(w / 6, h).Z();
    return { l: w / 6, t: 0, r: (w * 5) / 6, b: h };
  },
  flowChartMagneticDisk: (p, w, h) => {
    const y1 = h / 6;
    p.begin('norm', false).M(0, y1).arc(w / 2, y1, 180, -180).L(w, h - y1).arc(w / 2, y1, 0, 180).Z();
    p.begin('lighten', false).M(0, y1).arc(w / 2, y1, 180, 360).Z();
    p.begin('none').M(w, y1).arc(w / 2, y1, 0, 180).arc(w / 2, y1, 180, 180).L(w, h - y1).arc(w / 2, y1, 0, 180).L(0, y1);
    return { l: 0, t: y1 * 2, r: w, b: h - y1 };
  },
  flowChartSummingJunction: (p, w, h) => void p.ellipse(w / 2, h / 2, w / 2, h / 2),
  flowChartOr: (p, w, h) => {
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    p.begin('none').M(w / 2, 0).L(w / 2, h).M(0, h / 2).L(w, h / 2);
  },
  flowChartSort: (p, w, h) => {
    p.poly([[0, h / 2], [w / 2, 0], [w, h / 2], [w / 2, h]]);
    p.begin('none').M(0, h / 2).L(w, h / 2);
  },
  flowChartCollate: (p, w, h) => void p.poly([[0, 0], [w, 0], [0, h], [w, h]]),
  flowChartOnlineStorage: (p, w, h) => {
    p.M(w / 6, 0).L(w, 0).arc(w / 6, h / 2, 270, -180).L(w / 6, h).arc(w / 6, h / 2, 90, 180).Z();
    return { l: w / 6, t: 0, r: (w * 5) / 6, b: h };
  },
  can: (p, w, h, adj) => {
    const y1 = (ss(w, h) * clamp(adj('adj', 25000), 0, (50000 * h) / ss(w, h))) / 200000;
    p.begin('norm', false).M(0, y1).arc(w / 2, y1, 180, -180).L(w, h - y1).arc(w / 2, y1, 0, 180).Z();
    p.begin('lighten', false).M(0, y1).arc(w / 2, y1, 180, 360).Z();
    p.begin('none').M(w, y1).arc(w / 2, y1, 0, 180).arc(w / 2, y1, 180, 180).L(w, h - y1).arc(w / 2, y1, 0, 180).L(0, y1);
    return { l: 0, t: y1 * 2, r: w, b: h - y1 };
  },
  cube: (p, w, h, adj) => {
    const d = (ss(w, h) * clamp(adj('adj', 25000), 0, 100000)) / 100000;
    p.begin('norm', false).poly([[0, d], [w - d, d], [w - d, h], [0, h]]);
    p.begin('lightenLess', false).poly([[0, d], [d, 0], [w, 0], [w - d, d]]);
    p.begin('darkenLess', false).poly([[w - d, h], [w - d, d], [w, 0], [w, h - d]]);
    p.begin('none').poly([[0, d], [d, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]).M(0, d).L(w - d, d).L(w, 0).M(w - d, d).L(w - d, h);
    return { l: 0, t: d, r: w - d, b: h };
  },
  bevel: (p, w, h, adj) => {
    const d = (ss(w, h) * clamp(adj('adj', 12500), 0, 50000)) / 100000;
    p.begin('norm', false).poly([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]);
    p.begin('lightenLess', false).poly([[0, 0], [w, 0], [w - d, d], [d, d]]);
    p.begin('darkenLess', false).poly([[0, h], [d, h - d], [w - d, h - d], [w, h]]);
    p.begin('lighten', false).poly([[0, 0], [d, d], [d, h - d], [0, h]]);
    p.begin('darken', false).poly([[w, 0], [w, h], [w - d, h - d], [w - d, d]]);
    p.begin('none').poly([[0, 0], [w, 0], [w, h], [0, h]]).poly([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]);
    return { l: d, t: d, r: w - d, b: h - d };
  },
  wedgeRectCallout: (p, w, h, adj) => void wedgeCallout(p, w, h, adj, 0),
  wedgeRoundRectCallout: (p, w, h, adj) => {
    const r = (ss(w, h) * clamp(adj('adj3', 16667), 0, 50000)) / 100000;
    wedgeCallout(p, w, h, adj, r);
    return { l: r * 0.29289, t: r * 0.29289, r: w - r * 0.29289, b: h - r * 0.29289 };
  },
  wedgeEllipseCallout: (p, w, h, adj) => {
    const tx = w / 2 + (w * adj('adj1', -20833)) / 100000, ty = h / 2 + (h * adj('adj2', 62500)) / 100000;
    const angle = Math.atan2((ty - h / 2) / (h || 1), (tx - w / 2) / (w || 1)) / DEG;
    const at = (deg: number): [number, number] => [w / 2 + (w / 2) * Math.cos(deg * DEG), h / 2 + (h / 2) * Math.sin(deg * DEG)];
    const [sx, sy] = at(angle + 10);
    p.M(sx, sy).arc(w / 2, h / 2, angle + 10, 340).L(tx, ty).Z();
    return { l: w * ELLIPSE_INSET, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  cloud: (p, w, h) => {
    // An approximation: a ring of overlapping bumps instead of the exact outline.
    const bumps = 10;
    p.M(w / 2 + (w / 2) * 0.86, h / 2);
    for (let i = 0; i < bumps; i++) {
      const a0 = (360 / bumps) * i, a1 = a0 + 360 / bumps;
      const x1 = w / 2 + (w / 2) * 0.86 * Math.cos(a1 * DEG), y1 = h / 2 + (h / 2) * 0.86 * Math.sin(a1 * DEG);
      const mid = ((a0 + a1) / 2) * DEG;
      p.Q(w / 2 + (w / 2) * 1.16 * Math.cos(mid), h / 2 + (h / 2) * 1.16 * Math.sin(mid), x1, y1);
    }
    p.Z();
    return { l: w * 0.15, t: h * 0.2, r: w * 0.85, b: h * 0.8 };
  },
  borderCallout1: callout([18750, -8333, 112500, -38333], true, false),
  borderCallout2: callout([18750, -8333, 18750, -16667, 112500, -46667], true, false),
  borderCallout3: callout([18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333], true, false),
  callout1: callout([18750, -8333, 112500, -38333], false, false),
  callout2: callout([18750, -8333, 18750, -16667, 112500, -46667], false, false),
  callout3: callout([18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333], false, false),
  accentCallout1: callout([18750, -8333, 112500, -38333], false, true),
  accentCallout2: callout([18750, -8333, 18750, -16667, 112500, -46667], false, true),
  accentBorderCallout1: callout([18750, -8333, 112500, -38333], true, true),
  accentBorderCallout2: callout([18750, -8333, 18750, -16667, 112500, -46667], true, true),
  leftBracket: bracket,
  rightBracket: mirrored(bracket, true, false),
  leftBrace: brace,
  rightBrace: mirrored(brace, true, false),
  bracketPair: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 16667), 0, 50000)) / 100000;
    p.begin('norm', false);
    roundRect(p, w, h, x1);
    p.begin('none').M(x1, h).arc(x1, x1, 90, 90).L(0, x1).arc(x1, x1, 180, 90);
    p.M(w - x1, 0).arc(x1, x1, 270, 90).L(w, h - x1).arc(x1, x1, 0, 90);
    return { l: x1 * 0.29289, t: x1 * 0.29289, r: w - x1 * 0.29289, b: h - x1 * 0.29289 };
  },
  bracePair: (p, w, h, adj) => {
    const x1 = (ss(w, h) * clamp(adj('adj', 8333), 0, 25000)) / 100000;
    p.begin('none').M(x1 * 2, h).arc(x1, x1, 90, 90).L(x1, h / 2 + x1).arc(x1, x1, 0, -90).arc(x1, x1, 90, -90).L(x1, x1).arc(x1, x1, 180, 90);
    p.M(w - x1 * 2, 0).arc(x1, x1, 270, 90).L(w - x1, h / 2 - x1).arc(x1, x1, 180, -90).arc(x1, x1, 270, -90).L(w - x1, h - x1).arc(x1, x1, 0, 90);
    return { l: x1 * 2, t: x1, r: w - x1 * 2, b: h - x1 };
  },
  arc: (p, w, h, adj) => {
    const st = adj('adj1', 16200000) / 60000;
    let sw = adj('adj2', 0) / 60000 - st;
    if (sw <= 0) sw += 360;
    const x = w / 2 + (w / 2) * Math.cos(st * DEG), y = h / 2 + (h / 2) * Math.sin(st * DEG);
    p.begin('norm', false).M(x, y).arc(w / 2, h / 2, st, sw).L(w / 2, h / 2).Z();
    p.begin('none').M(x, y).arc(w / 2, h / 2, st, sw);
  },
  pie: (p, w, h, adj) => {
    const st = adj('adj1', 0) / 60000;
    let sw = adj('adj2', 16200000) / 60000 - st;
    if (sw <= 0) sw += 360;
    p.M(w / 2 + (w / 2) * Math.cos(st * DEG), h / 2 + (h / 2) * Math.sin(st * DEG)).arc(w / 2, h / 2, st, sw).L(w / 2, h / 2).Z();
  },
  chord: (p, w, h, adj) => {
    const st = adj('adj1', 2700000) / 60000;
    let sw = adj('adj2', 16200000) / 60000 - st;
    if (sw <= 0) sw += 360;
    p.M(w / 2 + (w / 2) * Math.cos(st * DEG), h / 2 + (h / 2) * Math.sin(st * DEG)).arc(w / 2, h / 2, st, sw).Z();
  },
  blockArc: (p, w, h, adj) => {
    const st = adj('adj1', 10800000) / 60000;
    let sw = adj('adj2', 0) / 60000 - st;
    if (sw <= 0) sw += 360;
    const dr = (ss(w, h) * clamp(adj('adj3', 25000), 0, 50000)) / 100000;
    const rx = w / 2, ry = h / 2, ix = rx - dr, iy = ry - dr;
    const end = (st + sw) * DEG;
    p.M(w / 2 + rx * Math.cos(st * DEG), h / 2 + ry * Math.sin(st * DEG)).arc(rx, ry, st, sw);
    p.L(w / 2 + ix * Math.cos(end), h / 2 + iy * Math.sin(end)).arc(ix, iy, st + sw, -sw).Z();
  },
  pieWedge: (p, w, h) => void p.M(0, h).arc(w, h, 180, 90).L(w, h).Z(),
  teardrop: (p, w, h, adj) => {
    const a = clamp(adj('adj', 100000), 0, 200000) / 100000;
    // The tip moves along the diagonal: at 100 % it is the top right corner.
    const tx = w / 2 + (w / 2) * a, ty = h / 2 - (h / 2) * a;
    p.M(0, h / 2).arc(w / 2, h / 2, 180, 90).Q((w / 2 + tx) / 2 + w / 8, ty, tx, ty).Q(w, (h / 2 + ty) / 2 - h / 8, w, h / 2).arc(w / 2, h / 2, 0, 180).Z();
    return { l: w * ELLIPSE_INSET, t: h * ELLIPSE_INSET, r: w * (1 - ELLIPSE_INSET), b: h * (1 - ELLIPSE_INSET) };
  },
  heart: (p, w, h) => {
    const dx1 = (w * 49) / 48, dx2 = (w * 10) / 48, y1 = -h / 3;
    p.M(w / 2, h / 4).C(w / 2 + dx2, y1, w / 2 + dx1, h / 4, w / 2, h).C(w / 2 - dx1, h / 4, w / 2 - dx2, y1, w / 2, h / 4).Z();
    return { l: w / 4, t: h / 4, r: (w * 3) / 4, b: (h * 3) / 4 };
  },
  lightningBolt: (p, w, h) => {
    const pts = [[8458, 0], [12860, 6080], [11210, 6797], [16577, 12007], [14767, 12877], [21600, 21600], [10012, 14915], [12222, 13987], [5022, 9705], [7602, 8382], [0, 3890]];
    p.poly(pts.map(([x, y]) => [(w * x) / 21600, (h * y) / 21600]));
  },
  sun: (p, w, h) => {
    p.ellipse(w / 2, h / 2, w / 4, h / 4);
    for (let i = 0; i < 8; i++) {
      const a = i * 45 * DEG, s = 8 * DEG;
      const tip: [number, number] = [w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)];
      const b1: [number, number] = [w / 2 + w * 0.33 * Math.cos(a - s), h / 2 + h * 0.33 * Math.sin(a - s)];
      const b2: [number, number] = [w / 2 + w * 0.33 * Math.cos(a + s), h / 2 + h * 0.33 * Math.sin(a + s)];
      p.poly([b1, tip, b2]);
    }
    return { l: w * 0.32, t: h * 0.32, r: w * 0.68, b: h * 0.68 };
  },
  moon: (p, w, h, adj) => {
    const a = clamp(adj('adj', 50000), 0, 87500) / 100000;
    p.M(w, h).arc(w, h / 2, 90, 180).arc(w * (1 - a), h / 2, 270, -180).Z();
  },
  smileyFace: (p, w, h, adj) => {
    p.ellipse(w / 2, h / 2, w / 2, h / 2);
    p.begin('darkenLess').ellipse(w * 0.35, h * 0.38, w * 0.05, h * 0.05).ellipse(w * 0.65, h * 0.38, w * 0.05, h * 0.05);
    const smile = (h * clamp(adj('adj', 4653), -4653, 4653)) / 100000;
    p.begin('none').M(w * 0.23, h * 0.7 - smile).Q(w / 2, h * 0.7 + smile * 3, w * 0.77, h * 0.7 - smile);
  },
  wave: (p, w, h, adj) => {
    const y1 = (h * clamp(adj('adj1', 12500), 0, 20000)) / 100000;
    const k = 10 / 3;
    p.M(0, y1).C(w / 3, y1 - y1 * k, (w * 2) / 3, y1 + y1 * k, w, y1).L(w, h - y1).C((w * 2) / 3, h - y1 + y1 * k, w / 3, h - y1 - y1 * k, 0, h - y1).Z();
    return { l: 0, t: y1 * 2, r: w, b: h - y1 * 2 };
  },
  doubleWave: (p, w, h, adj) => {
    const y1 = (h * clamp(adj('adj1', 6250), 0, 12500)) / 100000;
    const k = 10 / 3;
    p.M(0, y1).C(w / 6, y1 - y1 * k, w / 3, y1 + y1 * k, w / 2, y1).C((w * 2) / 3, y1 - y1 * k, (w * 5) / 6, y1 + y1 * k, w, y1);
    p.L(w, h - y1).C((w * 5) / 6, h - y1 + y1 * k, (w * 2) / 3, h - y1 - y1 * k, w / 2, h - y1).C(w / 3, h - y1 + y1 * k, w / 6, h - y1 - y1 * k, 0, h - y1).Z();
    return { l: 0, t: y1 * 2, r: w, b: h - y1 * 2 };
  },
  mathPlus: (p, w, h, adj) => {
    const t = (ss(w, h) * clamp(adj('adj1', 23520), 0, 73490)) / 200000;
    const ax = w * 0.37, ay = h * 0.37, cx = w / 2, cy = h / 2;
    const dx = Math.min(ax, ay + t);
    p.poly([[cx - dx, cy - t], [cx - t, cy - t], [cx - t, cy - dx], [cx + t, cy - dx], [cx + t, cy - t], [cx + dx, cy - t], [cx + dx, cy + t], [cx + t, cy + t], [cx + t, cy + dx], [cx - t, cy + dx], [cx - t, cy + t], [cx - dx, cy + t]]);
  },
  mathMinus: (p, w, h, adj) => {
    const t = (h * clamp(adj('adj1', 23520), 0, 100000)) / 200000;
    p.poly([[w * 0.13, h / 2 - t], [w * 0.87, h / 2 - t], [w * 0.87, h / 2 + t], [w * 0.13, h / 2 + t]]);
  },
  mathEqual: (p, w, h, adj) => {
    const t = (h * clamp(adj('adj1', 23520), 0, 36745)) / 100000;
    const gap = (h * clamp(adj('adj2', 11760), 0, 100000)) / 200000;
    p.poly([[w * 0.13, h / 2 - gap - t], [w * 0.87, h / 2 - gap - t], [w * 0.87, h / 2 - gap], [w * 0.13, h / 2 - gap]]);
    p.poly([[w * 0.13, h / 2 + gap], [w * 0.87, h / 2 + gap], [w * 0.87, h / 2 + gap + t], [w * 0.13, h / 2 + gap + t]]);
  },
  mathMultiply: (p, w, h, adj) => {
    const t = (ss(w, h) * clamp(adj('adj1', 23520), 0, 51965)) / 200000;
    const cx = w / 2, cy = h / 2, r = Math.min(w, h) * 0.36;
    const arm = (a: number): [number, number][] => {
      const c = Math.cos(a), s = Math.sin(a);
      return [[cx + c * r - s * t, cy + s * r + c * t], [cx + c * r + s * t, cy + s * r - c * t]];
    };
    const pts: [number, number][] = [];
    for (const a of [45, 135, 225, 315]) {
      const [p1, p2] = arm(a * DEG);
      pts.push(p2, p1);
      const mid = (a + 45) * DEG;
      pts.push([cx + Math.cos(mid) * t * Math.SQRT2, cy + Math.sin(mid) * t * Math.SQRT2]);
    }
    p.poly(pts);
  },
  mathDivide: (p, w, h, adj) => {
    const t = (h * clamp(adj('adj1', 23520), 0, 36745)) / 200000;
    p.poly([[w * 0.13, h / 2 - t], [w * 0.87, h / 2 - t], [w * 0.87, h / 2 + t], [w * 0.13, h / 2 + t]]);
    const r = Math.min(w, h) * 0.08;
    p.ellipse(w / 2, h * 0.24, r, r).ellipse(w / 2, h * 0.76, r, r);
  },
  mathNotEqual: (p, w, h, adj) => {
    const t = (h * clamp(adj('adj1', 23520), 0, 50000)) / 100000;
    const gap = (h * 11760) / 200000;
    p.poly([[w * 0.13, h / 2 - gap - t], [w * 0.87, h / 2 - gap - t], [w * 0.87, h / 2 - gap], [w * 0.13, h / 2 - gap]]);
    p.poly([[w * 0.13, h / 2 + gap], [w * 0.87, h / 2 + gap], [w * 0.87, h / 2 + gap + t], [w * 0.13, h / 2 + gap + t]]);
    p.poly([[w * 0.55, h * 0.1], [w * 0.62, h * 0.1], [w * 0.45, h * 0.9], [w * 0.38, h * 0.9]]);
  },
  bentArrow: (p, w, h, adj) => {
    const s = ss(w, h);
    const th = (s * clamp(adj('adj1', 25000), 0, 50000)) / 100000;
    const hw = (s * clamp(adj('adj2', 25000), 0, 50000)) / 100000;
    const hl = (s * clamp(adj('adj3', 25000), 0, 50000)) / 100000;
    const r = Math.min((s * clamp(adj('adj4', 43750), 0, 100000)) / 100000, w - hl, h - hw);
    const y = hw - th / 2; // top of the horizontal shaft
    const ri = Math.max(0, r - th);
    p.M(0, h).L(0, y + r).arc(r, r, 180, 90).L(w - hl, y).L(w - hl, 0).L(w, hw).L(w - hl, hw * 2).L(w - hl, y + th);
    p.L(th + ri, y + th);
    if (ri > 0) p.arc(ri, ri, 270, -90);
    p.L(th, h).Z();
  },
  uturnArrow: (p, w, h, adj) => {
    const s = ss(w, h);
    const th = (s * clamp(adj('adj1', 25000), 0, 25000)) / 100000;
    const hw = (s * clamp(adj('adj2', 25000), 0, 25000)) / 100000;
    const hl = (s * clamp(adj('adj3', 25000), 0, 50000)) / 100000;
    const yEnd = (h * clamp(adj('adj5', 75000), 0, 100000)) / 100000;
    const xr = w - hw + th / 2; // right edge of the right shaft
    const r = Math.min((s * clamp(adj('adj4', 43750), 0, 100000)) / 100000, xr / 2);
    const ri = Math.max(0, r - th);
    p.M(0, h).L(0, r).arc(r, r, 180, 90).L(xr - r, 0).arc(r, r, 270, 90).L(xr, yEnd - hl).L(w, yEnd - hl).L(xr - th / 2, yEnd);
    p.L(w - 2 * hw, yEnd - hl).L(xr - th, yEnd - hl).L(xr - th, th + ri);
    if (ri > 0) p.arc(ri, ri, 0, -90);
    p.L(th + ri, th);
    if (ri > 0) p.arc(ri, ri, 270, -90);
    p.L(th, h).Z();
  },
  circularArrow: (p, w, h, adj) => {
    const th = (ss(w, h) * clamp(adj('adj1', 12500), 0, 25000)) / 100000;
    const st = adj('adj4', 10800000) / 60000;
    let end = adj('adj3', 20457681) / 60000;
    if (end <= st) end += 360;
    const rx = w / 2 - th, ry = h / 2 - th;
    const at = (rX: number, rY: number, deg: number): [number, number] => [w / 2 + rX * Math.cos(deg * DEG), h / 2 + rY * Math.sin(deg * DEG)];
    const [x0, y0] = at(rx + th / 2, ry + th / 2, st);
    p.M(x0, y0).arc(rx + th / 2, ry + th / 2, st, end - st);
    const [hx1, hy1] = at(rx + th * 1.5, ry + th * 1.5, end);
    const [tx, ty] = at(rx, ry, end + 12);
    const [hx2, hy2] = at(rx - th * 1.5, ry - th * 1.5, end);
    p.L(hx1, hy1).L(tx, ty).L(hx2, hy2);
    const [xi, yi] = at(rx - th / 2, ry - th / 2, end);
    p.L(xi, yi).arc(rx - th / 2, ry - th / 2, end, -(end - st)).Z();
  },
  funnel: (p, w, h) => {
    p.M(0, h * 0.1).arc(w / 2, h * 0.1, 180, 180).L(w * 0.57, h).L(w * 0.43, h).Z();
  },
};

// Shapes that only differ in name or decoration from one we have.
const ALIASES: Record<string, string> = {
  actionButtonBlank: 'rect', actionButtonHome: 'rect', actionButtonHelp: 'rect', actionButtonInformation: 'rect',
  actionButtonForwardNext: 'rect', actionButtonBackPrevious: 'rect', actionButtonEnd: 'rect', actionButtonBeginning: 'rect',
  actionButtonReturn: 'rect', actionButtonDocument: 'rect', actionButtonSound: 'rect', actionButtonMovie: 'rect',
  flowChartMagneticDrum: 'can', flowChartMagneticTape: 'ellipse', borderCallout: 'borderCallout1',
  accentBorderCallout3: 'accentBorderCallout2', accentCallout3: 'accentCallout2', cloudCallout: 'cloud',
  ellipseRibbon: 'wave', ellipseRibbon2: 'wave', ribbon: 'rect', ribbon2: 'rect', verticalScroll: 'roundRect', horizontalScroll: 'roundRect',
  irregularSeal1: 'star16', irregularSeal2: 'star16', gear6: 'star12', gear9: 'star16', plaqueTabs: 'rect', squareTabs: 'rect',
  cornerTabs: 'rect', chartPlus: 'rect', chartStar: 'rect', chartX: 'rect', leftRightRibbon: 'leftRightArrow',
  rightArrowCallout: 'rightArrow', leftArrowCallout: 'leftArrow', upArrowCallout: 'upArrow', downArrowCallout: 'downArrow',
  leftRightArrowCallout: 'leftRightArrow', upDownArrowCallout: 'upDownArrow', quadArrow: 'plus', quadArrowCallout: 'plus',
  leftRightUpArrow: 'leftRightArrow', leftUpArrow: 'bentArrow', bentUpArrow: 'bentArrow', curvedRightArrow: 'rightArrow',
  curvedLeftArrow: 'leftArrow', curvedUpArrow: 'upArrow', curvedDownArrow: 'downArrow', swooshArrow: 'rightArrow',
  leftCircularArrow: 'circularArrow', leftRightCircularArrow: 'circularArrow', flowChartOfflineStorage: 'flowChartMerge',
  verticalScroll2: 'roundRect', doubleBracket: 'bracketPair',
};

function avList(el: Element | null): Map<string, number> {
  const out = new Map<string, number>();
  for (const gd of children(el, 'gd')) {
    const name = attr(gd, 'name');
    const match = /^\s*val\s+(-?\d+(?:\.\d+)?)\s*$/.exec(attr(gd, 'fmla') ?? '');
    if (name && match) out.set(name, Number(match[1]));
  }
  return out;
}

export function presetGeometry(prst: string, w: number, h: number, av: Map<string, number>): Geometry {
  const name = PRESETS[prst] ? prst : (ALIASES[prst] ?? 'rect');
  const p = new PathBuilder();
  // Presets use "adj" when they have one handle; some files write "adj1" for those anyway.
  const adj: Adj = (key, fallback) => av.get(key) ?? (key === 'adj' ? av.get('adj1') : key === 'adj1' ? av.get('adj') : undefined) ?? fallback;
  const text = (PRESETS[name] ?? PRESETS.rect)(p, w, h, adj) ?? null;
  const open = prst === 'line' || prst === 'lineInv' || prst === 'arc' || /Connector\d/.test(prst) || prst.startsWith('leftBracket') || prst.startsWith('rightBracket') || prst.endsWith('Brace');
  return { paths: p.paths, text, open };
}

// ---- Custom geometry ------------------------------------------------------------------------------------------

type Vars = Map<string, number>;

function builtins(w: number, h: number): Vars {
  const s = Math.min(w, h), l = Math.max(w, h);
  const v: Vars = new Map([
    ['w', w], ['h', h], ['l', 0], ['t', 0], ['r', w], ['b', h], ['hc', w / 2], ['vc', h / 2], ['ss', s], ['ls', l],
    ['cd2', 10800000], ['cd4', 5400000], ['cd8', 2700000], ['3cd4', 16200000], ['3cd8', 8100000], ['5cd8', 13500000], ['7cd8', 18900000],
  ]);
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 32]) {
    v.set(`wd${n}`, w / n);
    v.set(`hd${n}`, h / n);
  }
  for (const n of [2, 4, 6, 8, 16, 32]) v.set(`ssd${n}`, s / n);
  return v;
}

function evaluate(formula: string, vars: Vars): number {
  const [op, ...args] = formula.trim().split(/\s+/);
  const a = args.map((x) => {
    const n = Number(x);
    return Number.isFinite(n) ? n : (vars.get(x) ?? 0);
  });
  const [x = 0, y = 0, z = 0] = a;
  const rad = (v: number) => (v / 60000) * DEG;
  switch (op) {
    case 'val':
      return x;
    case '*/':
      return z ? (x * y) / z : 0;
    case '+-':
      return x + y - z;
    case '+/':
      return z ? (x + y) / z : 0;
    case '?:':
      return x > 0 ? y : z;
    case 'abs':
      return Math.abs(x);
    case 'at2':
      return (Math.atan2(y, x) / DEG) * 60000;
    case 'cat2':
      return x * Math.cos(Math.atan2(z, y));
    case 'sat2':
      return x * Math.sin(Math.atan2(z, y));
    case 'cos':
      return x * Math.cos(rad(y));
    case 'sin':
      return x * Math.sin(rad(y));
    case 'tan':
      return x * Math.tan(rad(y));
    case 'max':
      return Math.max(x, y);
    case 'min':
      return Math.min(x, y);
    case 'mod':
      return Math.sqrt(x * x + y * y + z * z);
    case 'pin':
      return y < x ? x : y > z ? z : y;
    case 'sqrt':
      return Math.sqrt(Math.max(0, x));
    default:
      return 0;
  }
}

function guides(el: Element | null, vars: Vars, limit = 512): void {
  for (const gd of children(el, 'gd').slice(0, limit)) {
    const name = attr(gd, 'name');
    if (!name) continue;
    const value = evaluate(attr(gd, 'fmla') ?? '', vars);
    vars.set(name, Number.isFinite(value) ? value : 0);
  }
}

/** Custom geometry; coordinates are in EMU (or the path's own w/h), the result in pixels of the box. */
export function customGeometry(el: Element, w: number, h: number, emuW: number, emuH: number): Geometry {
  const vars = builtins(emuW, emuH);
  guides(child(el, 'avLst'), vars);
  guides(child(el, 'gdLst'), vars);
  const num = (value: string | null) => {
    if (value === null) return 0;
    const n = Number(value);
    return Number.isFinite(n) ? n : (vars.get(value) ?? 0);
  };
  const p = new PathBuilder();
  for (const path of children(child(el, 'pathLst'), 'path')) {
    const pw = int(path, 'w'), ph = int(path, 'h');
    const sx = pw ? w / pw : emuW ? w / emuW : 0;
    const sy = ph ? h / ph : emuH ? h / emuH : 0;
    const fill = attr(path, 'fill') ?? 'norm';
    p.begin(fill, attr(path, 'stroke') !== '0' && attr(path, 'stroke') !== 'false');
    const pt = (e: Element | null): [number, number] => [num(attr(e, 'x')) * sx, num(attr(e, 'y')) * sy];
    for (const cmd of path.children) {
      const pts = children(cmd, 'pt');
      switch (cmd.localName) {
        case 'moveTo':
          p.M(...pt(pts[0]));
          break;
        case 'lnTo':
          p.L(...pt(pts[0]));
          break;
        case 'cubicBezTo':
          if (pts.length >= 3) p.C(...pt(pts[0]), ...pt(pts[1]), ...pt(pts[2]));
          break;
        case 'quadBezTo':
          if (pts.length >= 2) p.Q(...pt(pts[0]), ...pt(pts[1]));
          break;
        case 'arcTo':
          p.arc(num(attr(cmd, 'wR')) * sx, num(attr(cmd, 'hR')) * sy, num(attr(cmd, 'stAng')) / 60000, num(attr(cmd, 'swAng')) / 60000);
          break;
        case 'close':
          p.Z();
          break;
      }
    }
  }
  const rect = child(el, 'rect');
  const scaleX = emuW ? w / emuW : 0, scaleY = emuH ? h / emuH : 0;
  const text = rect ? { l: num(attr(rect, 'l')) * scaleX, t: num(attr(rect, 't')) * scaleY, r: num(attr(rect, 'r')) * scaleX, b: num(attr(rect, 'b')) * scaleY } : null;
  const open = p.paths.every((sp) => !sp.cmds.some((c) => c.t === 'Z'));
  return { paths: p.paths, text, open };
}

/** Geometry of a shape from its spPr: preset, custom, or a plain rectangle. */
export function geometryOf(spPr: Element | null, w: number, h: number, emuW: number, emuH: number): Geometry {
  const prst = child(spPr, 'prstGeom');
  if (prst) return presetGeometry(attr(prst, 'prst') ?? 'rect', w, h, avList(child(prst, 'avLst')));
  const cust = child(spPr, 'custGeom');
  if (cust) return customGeometry(cust, w, h, emuW, emuH);
  return presetGeometry('rect', w, h, new Map());
}

const f = (n: number) => (Math.round(n * 100) / 100).toString();

export function svgPath(cmds: Cmd[]): string {
  return cmds
    .map((c) => (c.t === 'Z' ? 'Z' : c.t === 'C' ? `C${f(c.x1)} ${f(c.y1)} ${f(c.x2)} ${f(c.y2)} ${f(c.x)} ${f(c.y)}` : `${c.t}${f(c.x)} ${f(c.y)}`))
    .join('');
}

/** Point and direction of travel at the start or end of a path, for arrowheads. */
export function pathEnd(cmds: Cmd[], atStart: boolean): { x: number; y: number; dx: number; dy: number } | null {
  const drawn = cmds.filter((c) => c.t !== 'Z') as Exclude<Cmd, { t: 'Z' }>[];
  if (drawn.length < 2) return null;
  const dir = (x0: number, y0: number, x1: number, y1: number) => {
    const len = Math.hypot(x1 - x0, y1 - y0);
    return len > 1e-6 ? { dx: (x1 - x0) / len, dy: (y1 - y0) / len } : null;
  };
  if (atStart) {
    const first = drawn[0];
    for (const next of drawn.slice(1)) {
      const candidates = next.t === 'C' ? [[next.x1, next.y1], [next.x2, next.y2], [next.x, next.y]] : [[next.x, next.y]];
      for (const [x, y] of candidates) {
        const d = dir(x, y, first.x, first.y);
        if (d) return { x: first.x, y: first.y, ...d };
      }
    }
    return null;
  }
  const last = drawn[drawn.length - 1];
  const points: [number, number][] = [];
  for (const c of drawn.slice(0, -1)) points.push([c.x, c.y]);
  if (last.t === 'C') points.push([last.x1, last.y1], [last.x2, last.y2]);
  for (let i = points.length - 1; i >= 0; i--) {
    const d = dir(points[i][0], points[i][1], last.x, last.y);
    if (d) return { x: last.x, y: last.y, ...d };
  }
  return null;
}

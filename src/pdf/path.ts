// SPDX-License-Identifier: MIT
// SVG path data to the segments a PDF path consists of: moves, lines, cubic curves and closes, all absolute.
// Quadratic curves and elliptical arcs become cubic curves.

export type Segment =
  | { k: 'M'; x: number; y: number }
  | { k: 'L'; x: number; y: number }
  | { k: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { k: 'Z' };

const ARGS: Record<string, number> = { m: 2, l: 2, h: 1, v: 1, c: 6, s: 4, q: 4, t: 2, a: 7, z: 0 };

/** Reads path data character by character; arc flags may be written without separators ("a1 1 0 011 1"). */
class Scanner {
  private i = 0;
  private static readonly NUMBER = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y;

  constructor(private readonly s: string) {}

  private skip(): void {
    while (this.i < this.s.length && /[\s,]/.test(this.s[this.i])) this.i++;
  }

  done(): boolean {
    this.skip();
    return this.i >= this.s.length;
  }

  command(): string | null {
    this.skip();
    const c = this.s[this.i];
    if (c && /[MmLlHhVvCcSsQqTtAaZz]/.test(c)) {
      this.i++;
      return c;
    }
    return null;
  }

  number(): number | null {
    this.skip();
    Scanner.NUMBER.lastIndex = this.i;
    const m = Scanner.NUMBER.exec(this.s);
    if (!m) return null;
    this.i = Scanner.NUMBER.lastIndex;
    return Number(m[0]);
  }

  flag(): number | null {
    this.skip();
    const c = this.s[this.i];
    if (c !== '0' && c !== '1') return null;
    this.i++;
    return Number(c);
  }
}

export function parsePath(d: string): Segment[] {
  const scan = new Scanner(d);
  const out: Segment[] = [];
  let x = 0, y = 0, startX = 0, startY = 0;
  // Reflection points for S and T.
  let lastC: [number, number] | null = null;
  let lastQ: [number, number] | null = null;
  let cmd = '';
  const cubic = (x1: number, y1: number, x2: number, y2: number, ex: number, ey: number) => {
    out.push({ k: 'C', x1, y1, x2, y2, x: ex, y: ey });
    x = ex;
    y = ey;
  };
  while (!scan.done()) {
    cmd = scan.command() ?? cmd;
    if (!cmd) return out;
    const lower = cmd.toLowerCase();
    const rel = cmd !== cmd.toUpperCase();
    if (lower === 'z') {
      out.push({ k: 'Z' });
      x = startX;
      y = startY;
      lastC = lastQ = null;
      // Z takes no numbers; what follows needs a command of its own.
      cmd = '';
      continue;
    }
    const a: number[] = [];
    for (let k = 0; k < ARGS[lower]; k++) {
      const v = lower === 'a' && (k === 3 || k === 4) ? scan.flag() : scan.number();
      if (v === null || !Number.isFinite(v)) return out;
      a.push(v);
    }
    const ox = rel ? x : 0, oy = rel ? y : 0;
    switch (lower) {
      case 'm':
        x = a[0] + ox;
        y = a[1] + oy;
        startX = x;
        startY = y;
        out.push({ k: 'M', x, y });
        // Further pairs after a move are lines.
        cmd = rel ? 'l' : 'L';
        lastC = lastQ = null;
        break;
      case 'l':
        x = a[0] + ox;
        y = a[1] + oy;
        out.push({ k: 'L', x, y });
        lastC = lastQ = null;
        break;
      case 'h':
        x = a[0] + ox;
        out.push({ k: 'L', x, y });
        lastC = lastQ = null;
        break;
      case 'v':
        y = a[0] + oy;
        out.push({ k: 'L', x, y });
        lastC = lastQ = null;
        break;
      case 'c':
        cubic(a[0] + ox, a[1] + oy, a[2] + ox, a[3] + oy, a[4] + ox, a[5] + oy);
        lastC = [a[2] + ox, a[3] + oy];
        lastQ = null;
        break;
      case 's': {
        const [rx, ry] = lastC ? [2 * x - lastC[0], 2 * y - lastC[1]] : [x, y];
        cubic(rx, ry, a[0] + ox, a[1] + oy, a[2] + ox, a[3] + oy);
        lastC = [a[0] + ox, a[1] + oy];
        lastQ = null;
        break;
      }
      case 'q':
      case 't': {
        // The control point: given (Q), mirrored from the last one (T after Q or T), or the current point.
        const reflected: [number, number] = lastQ ? [2 * x - lastQ[0], 2 * y - lastQ[1]] : [x, y];
        const qx: number = lower === 'q' ? a[0] + ox : reflected[0];
        const qy: number = lower === 'q' ? a[1] + oy : reflected[1];
        const ex: number = lower === 'q' ? a[2] + ox : a[0] + ox;
        const ey: number = lower === 'q' ? a[3] + oy : a[1] + oy;
        cubic(x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), ex + (2 / 3) * (qx - ex), ey + (2 / 3) * (qy - ey), ex, ey);
        lastQ = [qx, qy];
        lastC = null;
        break;
      }
      case 'a':
        for (const c of arcToCubics(x, y, a[0], a[1], a[2], a[3] !== 0, a[4] !== 0, a[5] + ox, a[6] + oy)) cubic(...c);
        lastC = lastQ = null;
        break;
    }
  }
  return out;
}

/** An elliptical arc (SVG endpoint parameterization) as cubic curves of at most 90° each. */
export function arcToCubics(x1: number, y1: number, rx: number, ry: number, angle: number, large: boolean, sweep: boolean, x2: number, y2: number): [number, number, number, number, number, number][] {
  if (x1 === x2 && y1 === y2) return [];
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  if (!rx || !ry) return [[x1, y1, x2, y2, x2, y2]];
  const phi = (angle * Math.PI) / 180;
  const cos = Math.cos(phi), sin = Math.sin(phi);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
  const xp = cos * dx + sin * dy, yp = -sin * dx + cos * dy;
  const lambda = (xp * xp) / (rx * rx) + (yp * yp) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const num = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp;
  const den = rx * rx * yp * yp + ry * ry * xp * xp;
  let coef = Math.sqrt(Math.max(0, num / den));
  if (large === sweep) coef = -coef;
  const cxp = (coef * rx * yp) / ry, cyp = (-coef * ry * xp) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angleOf = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const t1 = angleOf(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
  let dt = angleOf((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  else if (sweep && dt < 0) dt += 2 * Math.PI;
  const parts = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
  const step = dt / parts;
  const k = (4 / 3) * Math.tan(step / 4);
  const point = (t: number): [number, number] => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
  const deriv = (t: number): [number, number] => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
  const out: [number, number, number, number, number, number][] = [];
  let t = t1;
  for (let p = 0; p < parts; p++) {
    const [ax, ay] = point(t);
    const [bx, by] = point(t + step);
    const [dax, day] = deriv(t);
    const [dbx, dby] = deriv(t + step);
    out.push([ax + k * dax, ay + k * day, bx - k * dbx, by - k * dby, bx, by]);
    t += step;
  }
  // The last point exactly, so closes and joins meet.
  const last = out[out.length - 1];
  last[4] = x2;
  last[5] = y2;
  return out;
}

/** A rectangle with optional rounded corners. */
export function rectPath(x: number, y: number, w: number, h: number, rx = 0, ry = rx): Segment[] {
  rx = Math.min(Math.max(0, rx), w / 2);
  ry = Math.min(Math.max(0, ry), h / 2);
  if (!rx || !ry) return [{ k: 'M', x, y }, { k: 'L', x: x + w, y }, { k: 'L', x: x + w, y: y + h }, { k: 'L', x, y: y + h }, { k: 'Z' }];
  const k = 0.5522847498;
  const ox = rx * k, oy = ry * k;
  return [
    { k: 'M', x: x + rx, y },
    { k: 'L', x: x + w - rx, y },
    { k: 'C', x1: x + w - rx + ox, y1: y, x2: x + w, y2: y + ry - oy, x: x + w, y: y + ry },
    { k: 'L', x: x + w, y: y + h - ry },
    { k: 'C', x1: x + w, y1: y + h - ry + oy, x2: x + w - rx + ox, y2: y + h, x: x + w - rx, y: y + h },
    { k: 'L', x: x + rx, y: y + h },
    { k: 'C', x1: x + rx - ox, y1: y + h, x2: x, y2: y + h - ry + oy, x, y: y + h - ry },
    { k: 'L', x, y: y + ry },
    { k: 'C', x1: x, y1: y + ry - oy, x2: x + rx - ox, y2: y, x: x + rx, y },
    { k: 'Z' },
  ];
}

export function ellipsePath(cx: number, cy: number, rx: number, ry: number): Segment[] {
  return rectPath(cx - rx, cy - ry, 2 * rx, 2 * ry, rx, ry);
}

export function polyPath(points: string, close: boolean): Segment[] {
  const n = (points.match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? []).map(Number);
  const out: Segment[] = [];
  for (let i = 0; i + 1 < n.length; i += 2) out.push({ k: out.length ? 'L' : 'M', x: n[i], y: n[i + 1] });
  if (close && out.length) out.push({ k: 'Z' });
  return out;
}

// SPDX-License-Identifier: MIT
// CSS gradients (computed background-image values) as PDF shadings: slide backgrounds and table cells in
// PowerPoint, Word's gradient text fill. The geometry follows the CSS rules for a box: the gradient line of a
// linear gradient runs through the center at its angle and is long enough to reach the corners; a radial one is
// a circle or ellipse sized by keyword, made from a circle in a scaled space.
import { parseColor } from './color.js';
import type { GradientStop, Matrix, Painter } from './painter.js';

interface Area {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface CssGradient {
  kind: 'linear' | 'radial';
  coords: number[];
  stops: GradientStop[];
  /** Space the coordinates are in (an ellipse is a scaled circle). */
  matrix?: Matrix;
}

/** Top-level comma-separated parts of a CSS value (commas inside parentheses stay). */
export function splitTop(value: string): string[] {
  const out: string[] = [];
  let depth = 0, start = 0;
  for (let i = 0; i < value.length; i++) {
    const c = value[i];
    if (c === '(') depth++;
    else if (c === ')') depth--;
    else if (c === ',' && depth === 0) {
      out.push(value.slice(start, i).trim());
      start = i + 1;
    }
  }
  out.push(value.slice(start).trim());
  return out.filter(Boolean);
}

const COLOR_START = /^(rgba?\(|color\(|hsla?\(|#|transparent\b|currentcolor\b|[a-z]+$|[a-z]+\s)/i;

/** A length in a gradient (px or % of `basis`). */
function length(token: string, basis: number): number | null {
  const n = parseFloat(token);
  if (!Number.isFinite(n)) return null;
  return token.endsWith('%') ? (n / 100) * basis : n;
}

/** Color stops with positions as fractions of `basis` px; missing positions spread evenly, as CSS does. */
function stopsOf(parts: string[], basis: number, alpha: number): GradientStop[] {
  const raw: { color: GradientStop['color']; offset: number | null }[] = [];
  for (const part of parts) {
    const m = /^((?:rgba?|color|hsla?)\([^)]*\)|#[0-9a-f]+|[a-z]+)\s*(.*)$/i.exec(part);
    if (!m) continue; // a color hint (a bare position) changes the blend only slightly
    const color = parseColor(m[1], alpha, true);
    if (!color) continue;
    const positions = m[2].split(/\s+/).filter(Boolean).map((t) => length(t, basis));
    if (!positions.length) raw.push({ color, offset: null });
    for (const p of positions) raw.push({ color, offset: p === null ? null : p / (basis || 1) });
  }
  if (!raw.length) return [];
  raw[0].offset ??= 0;
  raw[raw.length - 1].offset ??= 1;
  // Positions never go back; gaps without one are filled in steps.
  let last = raw[0].offset;
  for (let i = 1; i < raw.length; i++) {
    if (raw[i].offset === null) {
      let j = i;
      while (raw[j].offset === null) j++;
      const end = Math.max(last, raw[j].offset!);
      for (let k = i; k < j; k++) raw[k].offset = last + ((end - last) * (k - i + 1)) / (j - i + 1);
    }
    raw[i].offset = Math.max(last, raw[i].offset!);
    last = raw[i].offset!;
  }
  return raw.map((s) => ({ color: s.color, offset: s.offset! }));
}

/** Unit direction of a linear gradient (y down) from its first argument, or null when it is a color stop. */
function direction(spec: string, area: Area): [number, number] | null {
  const angle = /^(-?[\d.]+)(deg|grad|rad|turn)$/.exec(spec);
  if (angle) {
    const v = parseFloat(angle[1]);
    const rad = angle[2] === 'deg' ? (v * Math.PI) / 180 : angle[2] === 'grad' ? (v * Math.PI) / 200 : angle[2] === 'turn' ? v * 2 * Math.PI : v;
    return [Math.sin(rad), -Math.cos(rad)];
  }
  const to = /^to\s+(.+)$/.exec(spec);
  if (!to) return null;
  const words = to[1].split(/\s+/);
  const sx = words.includes('right') ? 1 : words.includes('left') ? -1 : 0;
  const sy = words.includes('bottom') ? 1 : words.includes('top') ? -1 : 0;
  // Toward a corner the line stands at right angles to the diagonal through the other two corners.
  const dx = sx && sy ? sx * area.h : sx, dy = sx && sy ? sy * area.w : sy;
  const n = Math.hypot(dx, dy) || 1;
  return [dx / n, dy / n];
}

function linear(args: string[], area: Area, alpha: number): CssGradient | null {
  const dir = direction(args[0], area);
  const [dx, dy] = dir ?? [0, 1];
  const stops = dir ? args.slice(1) : args;
  const len = Math.abs(area.w * dx) + Math.abs(area.h * dy);
  const cx = area.x + area.w / 2, cy = area.y + area.h / 2;
  return {
    kind: 'linear',
    coords: [cx - (dx * len) / 2, cy - (dy * len) / 2, cx + (dx * len) / 2, cy + (dy * len) / 2],
    stops: stopsOf(stops, len, alpha),
  };
}

/** A position keyword or length along one axis. */
function axis(token: string | undefined, size: number): number {
  if (!token || token === 'center') return size / 2;
  if (token === 'left' || token === 'top') return 0;
  if (token === 'right' || token === 'bottom') return size;
  return length(token, size) ?? size / 2;
}

function radial(args: string[], area: Area, alpha: number): CssGradient | null {
  let spec = '';
  let stops = args;
  if (!COLOR_START.test(args[0]) || /^(circle|ellipse|closest|farthest|at)\b/.test(args[0])) {
    spec = args[0];
    stops = args.slice(1);
  }
  const [shapePart, atPart] = spec.split(/\bat\b/).map((s) => s.trim());
  const at = (atPart ?? '').split(/\s+/).filter(Boolean);
  // "at top" and "at left" name one axis; the other stays centered.
  if (at.length === 1 && (at[0] === 'top' || at[0] === 'bottom')) at.unshift('center');
  const cx = axis(at[0], area.w), cy = axis(at[1], area.h);
  const words = (shapePart ?? '').split(/\s+/).filter(Boolean);
  const lengths = words.filter((w) => /^-?[\d.]/.test(w));
  const circle = words.includes('circle') || (!words.includes('ellipse') && lengths.length === 1);
  const sides = [cx, area.w - cx], tops = [cy, area.h - cy];
  const keyword = words.find((w) => /^(closest|farthest)-(side|corner)$/.test(w)) ?? 'farthest-corner';
  const pick = keyword.startsWith('closest') ? Math.min : Math.max;
  let rx: number, ry: number;
  if (lengths.length) {
    rx = length(lengths[0], area.w) ?? 0;
    ry = circle ? rx : (length(lengths[1] ?? lengths[0], area.h) ?? rx);
  } else if (circle) {
    if (keyword.endsWith('side')) rx = pick(...sides, ...tops);
    else rx = pick(...sides.flatMap((x) => tops.map((y) => Math.hypot(x, y))));
    ry = rx;
  } else {
    rx = pick(...sides);
    ry = pick(...tops);
    // An ellipse to a corner keeps the side ellipse's proportions.
    if (keyword.endsWith('corner')) {
      rx *= Math.SQRT2;
      ry *= Math.SQRT2;
    }
  }
  if (rx <= 0 || ry <= 0) return null;
  return {
    kind: 'radial',
    coords: [0, 0, 0, 0, 0, rx],
    stops: stopsOf(stops, rx, alpha),
    matrix: [1, 0, 0, ry / rx, area.x + cx, area.y + cy],
  };
}

/** A computed gradient value for a box, or null for anything else (url(), repeating gradients). */
export function parseGradient(value: string, area: Area, alpha = 1): CssGradient | null {
  const m = /^(linear|radial)-gradient\((.*)\)$/s.exec(value.trim());
  if (!m) return null;
  const args = splitTop(m[2]);
  if (!args.length) return null;
  const g = m[1] === 'linear' ? linear(args, area, alpha) : radial(args, area, alpha);
  return g && g.stops.length ? g : null;
}

/** Fills the current clip with the gradient. */
export function paintGradient(painter: Painter, g: CssGradient, opacity = 1): void {
  painter.save();
  if (g.matrix) painter.transform(g.matrix);
  painter.gradient(g.kind, g.coords, g.stops, opacity);
  painter.restore();
}

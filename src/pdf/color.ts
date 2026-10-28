// SPDX-License-Identifier: MIT
// Computed CSS colors (rgb(), rgba(), color(srgb …)) as PDF colors with components from 0 to 1.
import type { Rgba } from './painter.js';

/** A color, or null for none; fully transparent colors count as none unless `keepClear` (gradient stops). */
export function parseColor(value: string | null | undefined, alpha = 1, keepClear = false): Rgba | null {
  if (!value) return null;
  const v = value.trim();
  if (keepClear && v === 'transparent') return { r: 0, g: 0, b: 0, a: 0 };
  if (!v || v === 'none' || v === 'transparent' || v === 'currentcolor') return null;
  let m = /^rgba?\(([^)]*)\)$/i.exec(v);
  if (m) {
    const parts = m[1].split(/[\s,/]+/).filter(Boolean);
    const channel = (s: string) => (s.endsWith('%') ? (parseFloat(s) / 100) * 255 : parseFloat(s));
    const [r, g, b] = parts.slice(0, 3).map(channel);
    const a = parts[3] === undefined ? 1 : parts[3].endsWith('%') ? parseFloat(parts[3]) / 100 : parseFloat(parts[3]);
    if (![r, g, b, a].every(Number.isFinite)) return null;
    const color = { r: r / 255, g: g / 255, b: b / 255, a: a * alpha };
    return color.a > 0 || keepClear ? color : null;
  }
  m = /^color\(srgb\s+([\d.e+-]+)\s+([\d.e+-]+)\s+([\d.e+-]+)(?:\s*\/\s*([\d.e+-]+%?))?\)$/i.exec(v);
  if (m) {
    const a = m[4] === undefined ? 1 : m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4]);
    return { r: +m[1], g: +m[2], b: +m[3], a: a * alpha };
  }
  if (/^#[0-9a-f]{3,8}$/i.test(v)) {
    const hex = v.length <= 5 ? [...v.slice(1)].map((c) => c + c).join('') : v.slice(1);
    const n = (i: number) => parseInt(hex.slice(i, i + 2), 16) / 255;
    return { r: n(0), g: n(2), b: n(4), a: (hex.length === 8 ? n(6) : 1) * alpha };
  }
  return null;
}

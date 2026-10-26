// SPDX-License-Identifier: MIT
// Charts (DrawingML chart parts) drawn as SVG from the values cached in the file: column/bar, line, area,
// scatter, pie and doughnut, with axes, gridlines, legend, title and simple data labels. Other chart types
// get a labeled placeholder. Linked workbooks are never opened.
import { type Color, colorOf, css } from './color.js';
import { type Env, type Source, r2, svg } from './env.js';
import { type Fill, type Line, fillColor, ownFill, paintFill, paintStroke, parseLine } from './fill.js';
import { fontFamily } from './fonts.js';
import type { Box } from './shapes.js';
import { attr, child, children, clamp, descendant, elements, int, parseXml, val } from './xml.js';

interface TextStyle {
  size: number;
  color: string;
  font: string;
  bold: boolean;
  italic: boolean;
  caps: boolean;
}

interface Series {
  name: string;
  values: (number | null)[];
  xValues: (number | null)[] | null;
  categories: string[];
  format: string;
  fill: Fill | undefined;
  line: Line | undefined;
  points: Map<number, Fill>;
  marker: { symbol: string; size: number; fill: Fill | undefined; line: Line | undefined };
  smooth: boolean;
  labels: Labels | null;
}

interface Labels {
  value: boolean;
  category: boolean;
  series: boolean;
  percent: boolean;
  pos: string;
  format: string | null;
  style: TextStyle;
}

interface Group {
  type: string;
  el: Element;
  series: Series[];
  axIds: string[];
}

interface Axis {
  id: string;
  kind: string;
  el: Element;
  pos: string;
  deleted: boolean;
  style: TextStyle;
  format: string | null;
  sourceLinked: boolean;
  line: Line | undefined;
  grid: Line | undefined;
  min?: number;
  max?: number;
  major?: number;
  reversed: boolean;
  title: string | null;
  titleStyle: TextStyle;
  crossBetween: string;
}

const MAX_POINTS = 5000;
const MAX_SERIES = 256;
const PT = 4 / 3;
const NOWHERE: Source = { part: '', rels: new Map() };

let measureContext: CanvasRenderingContext2D | null | undefined;

function measure(text: string, style: TextStyle): number {
  if (measureContext === undefined) measureContext = document.createElement('canvas').getContext('2d');
  if (!measureContext) return text.length * style.size * 0.55;
  measureContext.font = `${style.italic ? 'italic ' : ''}${style.bold ? '700 ' : ''}${style.size}px ${style.font}`;
  return measureContext.measureText(text).width;
}

const MUTED = parseXml('<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>');

function baseStyle(env: Env): TextStyle {
  const muted = MUTED ? colorOf(MUTED.documentElement, env.colors) : null;
  return { size: 10 * PT, color: muted ? css(muted) : '#595959', font: fontFamily(env.theme.minor.latin), bold: false, italic: false, caps: false };
}

function themeFont(name: string, env: Env): string {
  if (name.startsWith('+mj')) return env.theme.major.latin;
  if (name.startsWith('+mn')) return env.theme.minor.latin;
  return name;
}

/** Text style from a run property element (defRPr, rPr) on top of a base style. */
function runStyle(rPr: Element | null, base: TextStyle, env: Env): TextStyle {
  if (!rPr) return base;
  const style = { ...base };
  const sz = int(rPr, 'sz');
  if (sz !== undefined) style.size = (clamp(sz, 100, 40000) / 100) * PT;
  const b = attr(rPr, 'b'), i = attr(rPr, 'i');
  if (b !== null) style.bold = b === '1' || b === 'true';
  if (i !== null) style.italic = i === '1' || i === 'true';
  const cap = attr(rPr, 'cap');
  if (cap !== null) style.caps = cap === 'all';
  const fill = ownFill(rPr, env.colors, NOWHERE);
  const color = fillColor(fill);
  if (color) style.color = css(color);
  const latin = attr(child(rPr, 'latin'), 'typeface');
  if (latin) style.font = fontFamily(themeFont(latin, env));
  return style;
}

/** Style given by a txPr (or rich) element: the first paragraph's default run properties. */
function textStyle(container: Element | null, base: TextStyle, env: Env): TextStyle {
  const p = child(container, 'p');
  return runStyle(child(child(p, 'pPr'), 'defRPr'), base, env);
}

/** Text of a title or label; runs set to all caps (cap="all", also inherited from the paragraph) are uppercased. */
function richText(tx: Element | null): string | null {
  if (!tx) return null;
  const rich = child(tx, 'rich');
  if (rich) {
    return children(rich, 'p')
      .map((p) => {
        const caps = attr(child(child(p, 'pPr'), 'defRPr'), 'cap');
        return children(p, 'r')
          .map((r) => {
            const text = child(r, 't')?.textContent ?? '';
            return (attr(child(r, 'rPr'), 'cap') ?? caps) === 'all' ? text.toUpperCase() : text;
          })
          .join('');
      })
      .join('\n');
  }
  const cached = descendant(tx, 'v');
  return cached?.textContent ?? null;
}

// ---- Data ---------------------------------------------------------------------------------------------------

function cache(ref: Element | null): { values: (string | null)[]; format: string } {
  const data = ref ? (descendant(ref, 'numCache') ?? descendant(ref, 'strCache') ?? descendant(ref, 'numLit') ?? descendant(ref, 'strLit') ?? descendant(ref, 'lvl')) : null;
  const count = clamp(int(child(data, 'ptCount'), 'val') ?? 0, 0, MAX_POINTS);
  const values: (string | null)[] = new Array(count).fill(null);
  for (const pt of children(data, 'pt')) {
    const idx = int(pt, 'idx') ?? -1;
    if (idx < 0 || idx >= MAX_POINTS) continue;
    while (values.length <= idx) values.push(null);
    values[idx] = child(pt, 'v')?.textContent ?? null;
  }
  return { values, format: child(data, 'formatCode')?.textContent ?? 'General' };
}

function numbers(values: (string | null)[]): (number | null)[] {
  return values.map((v) => {
    if (v === null || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  });
}

function parseLabels(el: Element | null, base: TextStyle, env: Env): Labels | null {
  if (!el || val(el, 'delete') === '1') return null;
  const on = (name: string) => val(el, name) === '1';
  const labels: Labels = {
    value: on('showVal'),
    category: on('showCatName'),
    series: on('showSerName'),
    percent: on('showPercent'),
    pos: val(el, 'dLblPos') ?? '',
    format: attr(child(el, 'numFmt'), 'sourceLinked') === '1' ? null : attr(child(el, 'numFmt'), 'formatCode'),
    style: textStyle(child(el, 'txPr'), base, env),
  };
  return labels.value || labels.category || labels.series || labels.percent ? labels : null;
}

function parseSeries(ser: Element, group: Element, base: TextStyle, env: Env): Series {
  const spPr = child(ser, 'spPr');
  const data = cache(child(ser, 'val') ?? child(ser, 'yVal'));
  const cat = cache(child(ser, 'cat') ?? child(ser, 'xVal'));
  const xVal = child(ser, 'xVal');
  const points = new Map<number, Fill>();
  for (const dPt of children(ser, 'dPt')) {
    const fill = ownFill(child(dPt, 'spPr'), env.colors, NOWHERE);
    const idx = int(child(dPt, 'idx'), 'val');
    if (fill && idx !== undefined) points.set(idx, fill);
  }
  const marker = child(ser, 'marker');
  const markerSpPr = child(marker, 'spPr');
  return {
    name: richText(child(ser, 'tx')) ?? `Series ${(int(child(ser, 'idx'), 'val') ?? 0) + 1}`,
    values: numbers(data.values),
    xValues: xVal ? numbers(cat.values) : null,
    categories: cat.values.map((v) => v ?? ''),
    format: data.format,
    fill: ownFill(spPr, env.colors, NOWHERE),
    line: parseLine(child(spPr, 'ln'), env.colors, NOWHERE),
    points,
    marker: {
      symbol: val(marker, 'symbol') ?? (group.localName === 'lineChart' && val(group, 'marker') !== '0' ? 'auto' : group.localName === 'scatterChart' ? 'auto' : 'none'),
      size: clamp(int(child(marker, 'size'), 'val') ?? 5, 2, 72) * PT,
      fill: ownFill(markerSpPr, env.colors, NOWHERE),
      line: parseLine(child(markerSpPr, 'ln'), env.colors, NOWHERE),
    },
    smooth: val(ser, 'smooth') === '1',
    // Series labels replace the chart type's; a deleted series label means none.
    labels: child(ser, 'dLbls') ? parseLabels(child(ser, 'dLbls'), base, env) : parseLabels(child(group, 'dLbls'), base, env),
  };
}

// ---- Number formats -------------------------------------------------------------------------------------------

/** A practical subset of spreadsheet number formats: decimals, grouping, percent, literal text. */
export function formatNumber(value: number, code: string | null, locale: string): string {
  if (!code || /^general$/i.test(code.trim())) {
    return new Intl.NumberFormat(locale, { maximumSignificantDigits: 10, useGrouping: false }).format(value);
  }
  const sections = code.split(';');
  let section = sections[0];
  let v = value;
  if (value < 0 && sections[1] !== undefined) {
    section = sections[1];
    v = Math.abs(value);
  }
  // Drop colors and conditions ([Red], [>=100]) and turn escaped and quoted text into literals.
  section = section.replace(/\[[^\]]*\]/g, '').replace(/\\(.)/g, '"$1"').replace(/_./g, ' ').replace(/\*./g, '');
  const match = /[#0?][#0?,.\s]*%?/.exec(section.replace(/"[^"]*"/g, (m) => 'x'.repeat(m.length)));
  if (!match) return section.replace(/"/g, '');
  const pattern = match[0];
  const before = section.slice(0, match.index).replace(/"/g, '');
  const after = section.slice(match.index + pattern.length).replace(/"/g, '');
  const percent = pattern.includes('%') || after.includes('%');
  const [intPart, frac = ''] = pattern.replace('%', '').split('.');
  const minFrac = (frac.match(/0/g) ?? []).length;
  const maxFrac = minFrac + (frac.match(/[#?]/g) ?? []).length;
  const grouping = intPart.includes(',') && /[#0],[#0]/.test(intPart);
  const formatted = new Intl.NumberFormat(locale, { minimumFractionDigits: minFrac, maximumFractionDigits: maxFrac, useGrouping: grouping }).format(percent ? v * 100 : v);
  return `${before}${formatted}${percent && !after.includes('%') ? '%' : ''}${after}`;
}

// ---- Scales ---------------------------------------------------------------------------------------------------

function niceStep(range: number): number {
  const rough = range / 10;
  if (!(rough > 0)) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  for (const m of [1, 2, 5, 10]) if (m * magnitude >= rough * 0.999) return m * magnitude;
  return 10 * magnitude;
}

interface Scale {
  min: number;
  max: number;
  step: number;
}

/** Next step in the 1, 2, 5 × 10ⁿ series. */
function nextStep(step: number): number {
  const magnitude = 10 ** Math.floor(Math.log10(step) + 1e-9);
  const m = Math.round(step / magnitude);
  return (m < 2 ? 2 : m < 5 ? 5 : 10) * magnitude;
}

/**
 * Automatic value axis like Excel's: zero-based unless the data sits far from zero, about ten intervals at
 * most, fewer where the axis is too short to space its labels (length in px, label font size in px).
 */
function valueScale(values: number[], axis: Axis | undefined, percentStacked: boolean, length = Infinity, label = 13): Scale {
  if (percentStacked) return { min: axis?.min ?? (Math.min(0, ...values) < 0 ? -1 : 0), max: axis?.max ?? 1, step: axis?.major ?? 0.1 };
  let lo = Math.min(...values), hi = Math.max(...values);
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) (lo = 0), (hi = 1);
  if (lo > 0 && (hi - lo) / hi > 1 / 6) lo = 0;
  if (hi < 0 && (lo - hi) / lo > 1 / 6) hi = 0;
  if (lo === hi) hi = lo + 1;
  const bounds = (step: number) => {
    let min = axis?.min ?? Math.floor(lo / step) * step;
    let max = axis?.max ?? Math.ceil(hi / step) * step;
    if (axis?.max === undefined && max - hi < (max - min) * 0.05 && hi !== 0) max += step;
    if (axis?.min === undefined && lo < 0 && lo - min < (max - min) * 0.05) min -= step;
    if (max <= min) max = min + step;
    return { min, max, step };
  };
  let scale = bounds(axis?.major ?? niceStep(hi - lo));
  if (!axis?.major) {
    for (let i = 0; i < 6 && ((scale.max - scale.min) / scale.step) * label * 1.9 > length; i++) scale = bounds(nextStep(scale.step));
  }
  return scale;
}

// ---- Drawing ------------------------------------------------------------------------------------------------

function textEl(parent: SVGElement, value: string, x: number, y: number, style: TextStyle, anchor: 'start' | 'middle' | 'end', rotate = 0): void {
  const t = svg('text', { x: r2(x), y: r2(y), 'text-anchor': anchor, 'dominant-baseline': 'central', fill: style.color, 'font-size': r2(style.size) });
  t.style.fontFamily = style.font;
  if (style.bold) t.setAttribute('font-weight', '700');
  if (style.italic) t.setAttribute('font-style', 'italic');
  if (rotate) t.setAttribute('transform', `rotate(${rotate} ${r2(x)} ${r2(y)})`);
  t.textContent = style.caps ? value.toUpperCase() : value;
  parent.append(t);
}

function rect(parent: SVGElement, x: number, y: number, w: number, h: number, fill: Fill | undefined, line: Line | undefined, env: Env, defs: () => SVGDefsElement): void {
  // Drawn at the origin of a moved group: gradients and picture fills are laid out in local coordinates.
  const el = svg('rect', { width: r2(Math.max(0, w)), height: r2(Math.max(0, h)) });
  const group = svg('g', { transform: `translate(${r2(x)} ${r2(y)})` });
  paintFill(el, fill, defs, w, h, env);
  paintStroke(el, line, defs, w, h);
  group.append(el);
  parent.append(group);
}

function lineEl(parent: SVGElement, x1: number, y1: number, x2: number, y2: number, line: Line | undefined, defs: () => SVGDefsElement): void {
  if (!line || line.fill.type === 'none') return;
  const el = svg('line', { x1: r2(x1), y1: r2(y1), x2: r2(x2), y2: r2(y2) });
  paintStroke(el, line, defs, 1, 1);
  parent.append(el);
}

function accent(index: number, env: Env): Color {
  const base = env.theme.colors.get(`accent${(index % 6) + 1}`) ?? { r: 68, g: 114, b: 196, a: 1 };
  const round = Math.floor(index / 6);
  if (!round) return base;
  // Later rounds of the palette are darker and lighter variants, like Office's automatic colors.
  const f = round % 2 ? 0.6 : 1.4;
  const c = (v: number) => clamp(f < 1 ? v * f : v + (255 - v) * (f - 1), 0, 255);
  return { r: c(base.r), g: c(base.g), b: c(base.b), a: 1 };
}

// Steps of a single-color palette: darker, lighter, the color itself, then further out.
const MONO_STEPS: [number, number][] = [[0, 0.75], [0.42, 1], [0, 1], [0, 0.55], [0.65, 1], [0.2, 1]];

/**
 * Automatic series colors. Charts with a legacy style (c:style 1–48) can use one color in shades:
 * gray for styles 1, 9, 17 …, an accent for 3–8, 11–16 …; style 2 and newer charts cycle the accents.
 */
function autoColor(index: number, style: number, env: Env): Color {
  const scheme = style > 0 && style <= 48 ? (style - 1) % 8 : 1;
  if (scheme === 1) return accent(index, env);
  const base = scheme === 0 ? { r: 136, g: 136, b: 136, a: 1 } : (env.theme.colors.get(`accent${scheme - 1}`) ?? { r: 68, g: 114, b: 196, a: 1 });
  const [toWhite, scale] = MONO_STEPS[index % MONO_STEPS.length];
  const c = (v: number) => clamp(v * scale + (255 - v * scale) * toWhite, 0, 255);
  return { r: c(base.r), g: c(base.g), b: c(base.b), a: 1 };
}

function solid(color: Color): Fill {
  return { type: 'solid', color };
}

function smoothPath(points: [number, number][]): string {
  if (points.length < 3) return points.map(([x, y], i) => `${i ? 'L' : 'M'}${r2(x)} ${r2(y)}`).join('');
  let d = `M${r2(points[0][0])} ${r2(points[0][1])}`;
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[Math.max(0, i - 1)], p1 = points[i], p2 = points[i + 1], p3 = points[Math.min(points.length - 1, i + 2)];
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6], c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += `C${r2(c1[0])} ${r2(c1[1])} ${r2(c2[0])} ${r2(c2[1])} ${r2(p2[0])} ${r2(p2[1])}`;
  }
  return d;
}

const AUTO_MARKERS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dash'];

function markerEl(parent: SVGElement, symbol: string, x: number, y: number, size: number, fill: Fill | undefined, line: Line | undefined, env: Env, defs: () => SVGDefsElement): void {
  const s = size / 2;
  let el: SVGElement;
  switch (symbol) {
    case 'circle':
      el = svg('circle', { cx: r2(x), cy: r2(y), r: r2(s) });
      break;
    case 'square':
      el = svg('rect', { x: r2(x - s), y: r2(y - s), width: r2(size), height: r2(size) });
      break;
    case 'triangle':
      el = svg('polygon', { points: `${r2(x)},${r2(y - s)} ${r2(x + s)},${r2(y + s)} ${r2(x - s)},${r2(y + s)}` });
      break;
    case 'x':
    case 'star':
    case 'plus':
      el = svg('path', { d: symbol === 'plus' ? `M${r2(x - s)} ${r2(y)}H${r2(x + s)}M${r2(x)} ${r2(y - s)}V${r2(y + s)}` : `M${r2(x - s)} ${r2(y - s)}L${r2(x + s)} ${r2(y + s)}M${r2(x - s)} ${r2(y + s)}L${r2(x + s)} ${r2(y - s)}` });
      break;
    case 'dash':
    case 'dot':
      el = svg('rect', { x: r2(x - s), y: r2(y - (symbol === 'dot' ? s / 2 : s / 4)), width: r2(size), height: r2(symbol === 'dot' ? s : s / 2) });
      break;
    default:
      el = svg('polygon', { points: `${r2(x)},${r2(y - s)} ${r2(x + s)},${r2(y)} ${r2(x)},${r2(y + s)} ${r2(x - s)},${r2(y)}` });
  }
  paintFill(el, fill, defs, size, size, env);
  paintStroke(el, line, defs, size, size);
  parent.append(el);
}

function parseAxis(el: Element, base: TextStyle, env: Env): Axis {
  const scaling = child(el, 'scaling');
  const num = (name: string, from: Element | null = scaling) => {
    const v = val(from, name);
    const n = v === null ? NaN : Number(v);
    return Number.isFinite(n) ? n : undefined;
  };
  const numFmt = child(el, 'numFmt');
  const spPr = child(el, 'spPr');
  const gridSpPr = child(child(el, 'majorGridlines'), 'spPr');
  const titleEl = child(el, 'title');
  const style = textStyle(child(el, 'txPr'), base, env);
  const major = num('majorUnit', el);
  return {
    id: val(el, 'axId') ?? '',
    kind: el.localName,
    el,
    pos: val(el, 'axPos') ?? 'b',
    deleted: val(el, 'delete') === '1',
    style,
    format: attr(numFmt, 'formatCode'),
    sourceLinked: attr(numFmt, 'sourceLinked') === '1',
    line: spPr ? parseLine(child(spPr, 'ln'), env.colors, NOWHERE) : { fill: solid({ r: 217, g: 217, b: 217, a: 1 }), width: 1, dash: 'solid', cap: 'butt', join: 'round' },
    grid: child(el, 'majorGridlines') ? (parseLine(child(gridSpPr, 'ln'), env.colors, NOWHERE) ?? { fill: solid({ r: 217, g: 217, b: 217, a: 1 }), width: 1, dash: 'solid', cap: 'butt', join: 'round' }) : undefined,
    min: num('min'),
    max: num('max'),
    major: major && major > 0 ? major : undefined,
    reversed: val(scaling, 'orientation') === 'maxMin',
    title: titleEl ? richText(child(titleEl, 'tx')) : null,
    titleStyle: { ...textStyle(child(titleEl, 'txPr') ?? child(child(titleEl, 'tx'), 'rich'), { ...style, bold: true }, env), caps: false },
    crossBetween: val(el, 'crossBetween') ?? 'between',
  };
}

const PLOTS = new Set(['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'area3DChart', 'scatterChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'ofPieChart', 'radarChart', 'bubbleChart', 'stockChart', 'surfaceChart', 'surface3DChart']);

export function renderChart(space: Element, container: HTMLElement, box: Box, env: Env): void {
  const W = Math.max(box.w, 1), H = Math.max(box.h, 1);
  const root = svg('svg', { class: 'pptx-chart-svg', width: r2(W), height: r2(H), viewBox: `0 0 ${r2(W)} ${r2(H)}` });
  let defs: SVGDefsElement | null = null;
  const getDefs = () => (defs ??= root.insertBefore(svg('defs'), root.firstChild));
  container.append(root);

  const chart = child(space, 'chart');
  const style = Number(val(space, 'style') ?? 0);
  const palette = (index: number) => autoColor(index, style, env);
  // Numbers follow the chart's own language (as PowerPoint does), English when it has none.
  const locale = val(space, 'lang') ?? 'en-US';
  const base = textStyle(child(space, 'txPr'), baseStyle(env), env);

  // Chart area.
  const spaceSpPr = child(space, 'spPr');
  rect(root, 0, 0, W, H, ownFill(spaceSpPr, env.colors, NOWHERE) ?? (spaceSpPr ? undefined : solid({ r: 255, g: 255, b: 255, a: 1 })), parseLine(child(spaceSpPr, 'ln'), env.colors, NOWHERE), env, getDefs);

  const plotArea = child(chart, 'plotArea');
  const groups: Group[] = [];
  for (const el of elements(plotArea)) {
    if (!PLOTS.has(el.localName)) continue;
    const series = children(el, 'ser').slice(0, MAX_SERIES).map((s) => parseSeries(s, el, base, env));
    groups.push({ type: el.localName, el, series, axIds: children(el, 'axId').map((a) => attr(a, 'val') ?? '') });
  }
  const axes = elements(plotArea).filter((e) => /^(catAx|valAx|dateAx|serAx)$/.test(e.localName)).map((e) => parseAxis(e, base, env));
  const allSeries = groups.flatMap((g) => g.series);

  // Title.
  let top = 8;
  const titleEl = child(chart, 'title');
  const titleDeleted = val(chart, 'autoTitleDeleted') === '1';
  let title = titleEl ? richText(child(titleEl, 'tx')) : null;
  if (titleEl && title === null && !titleDeleted && allSeries.length === 1) title = allSeries[0].name;
  if (title) {
    const style = { ...textStyle(child(titleEl, 'txPr') ?? child(child(titleEl, 'tx'), 'rich'), { ...base, size: 14 * PT }, env), caps: false };
    const lines = title.split('\n');
    lines.forEach((line, i) => textEl(root, line, W / 2, top + style.size * (0.6 + i * 1.2), style, 'middle'));
    if (val(titleEl, 'overlay') !== '1') top += style.size * 1.2 * lines.length + 6;
  }

  // Legend.
  const legendEl = child(chart, 'legend');
  const legendPos = val(legendEl, 'legendPos') ?? 'r';
  const legendStyle = textStyle(child(legendEl, 'txPr'), base, env);
  const isPie = groups.length > 0 && /pie|doughnut/i.test(groups[0].type);
  const deletedEntries = new Set(children(legendEl, 'legendEntry').filter((e) => val(e, 'delete') === '1').map((e) => int(child(e, 'idx'), 'val')));
  const entries: { label: string; fill: Fill | undefined; line: Line | undefined; marker: boolean }[] = [];
  if (legendEl) {
    if (isPie) {
      const s = groups[0].series[0];
      s?.values.forEach((_, i) => {
        if (!deletedEntries.has(i)) entries.push({ label: s.categories[i] ?? String(i + 1), fill: s.points.get(i) ?? solid(palette(i)), line: undefined, marker: false });
      });
    } else {
      let i = 0;
      for (const g of groups) {
        for (const s of g.series) {
          const index = i++;
          if (deletedEntries.has(index)) continue;
          const lineLike = /line|scatter|radar/i.test(g.type);
          entries.push({ label: s.name, fill: lineLike ? undefined : (s.fill ?? solid(palette(index))), line: lineLike ? (s.line ?? { fill: solid(palette(index)), width: 2, dash: 'solid', cap: 'round', join: 'round' }) : undefined, marker: lineLike });
        }
      }
    }
  }
  let left = 8, right = W - 8, bottom = H - 8;
  if (entries.length) {
    const sw = legendStyle.size * 0.8;
    const widths = entries.map((e) => measure(e.label, legendStyle) + sw + 6);
    const rowH = legendStyle.size * 1.45;
    const drawEntry = (e: (typeof entries)[0], x: number, y: number) => {
      if (e.line) lineEl(root, x, y, x + sw, y, { ...e.line, width: Math.min(e.line.width, 3) }, getDefs);
      else rect(root, x, y - sw / 2, sw, sw, e.fill, undefined, env, getDefs);
      textEl(root, e.label, x + sw + 4, y, legendStyle, 'start');
    };
    if (legendPos === 'r' || legendPos === 'l' || legendPos === 'tr') {
      const colW = Math.min(Math.max(...widths) + 8, W * 0.4);
      const x = legendPos === 'l' ? left : right - colW;
      const startY = (top + bottom) / 2 - (entries.length * rowH) / 2 + rowH / 2;
      entries.forEach((e, i) => drawEntry(e, x, startY + i * rowH));
      if (legendPos === 'l') left += colW + 6;
      else right -= colW + 6;
    } else {
      // One centered row when everything fits, otherwise a grid of equal columns filled row by row,
      // like Office lays out legends.
      const total = widths.reduce((a, w) => a + w + 12, -12);
      const single = total <= W - 16;
      const colW = Math.min(Math.max(...widths) + 12, W - 16);
      const cols = single ? entries.length : Math.max(1, Math.min(entries.length, Math.floor((W - 16) / colW)));
      const rows = Math.ceil(entries.length / cols);
      const blockH = rows * rowH;
      const y0 = legendPos === 't' ? top + rowH / 2 : bottom - blockH + rowH / 2;
      if (single) {
        let x = W / 2 - total / 2;
        entries.forEach((e, i) => {
          drawEntry(e, x, y0);
          x += widths[i] + 12;
        });
      } else {
        const x0 = W / 2 - (cols * colW - 12) / 2;
        entries.forEach((e, i) => drawEntry(e, x0 + (i % cols) * colW, y0 + Math.floor(i / cols) * rowH));
      }
      if (legendPos === 't') top += blockH + 8;
      else bottom -= blockH + 8;
    }
  }

  if (!groups.length || groups.some((g) => /radar|bubble|stock|surface|ofPie/.test(g.type))) {
    // Chart types we do not draw get a neutral frame with a label instead of something misleading.
    rect(root, left, top, right - left, bottom - top, undefined, { fill: solid({ r: 160, g: 160, b: 160, a: 1 }), width: 1, dash: 'dash', cap: 'butt', join: 'round' }, env, getDefs);
    textEl(root, 'Chart', (left + right) / 2, (top + bottom) / 2, base, 'middle');
    return;
  }

  // Plot area: a manual layout wins, otherwise what is left.
  const manual = descendant(child(plotArea, 'layout'), 'manualLayout');
  let plot = { x: left, y: top, w: right - left, h: bottom - top };
  let manualInner = false;
  if (manual) {
    const f = (name: string) => Number(val(manual, name) ?? NaN);
    const [x, y, w, h] = [f('x'), f('y'), f('w'), f('h')];
    if ([x, y, w, h].every(Number.isFinite) && (val(manual, 'xMode') ?? 'factor') === 'edge') {
      plot = { x: x * W, y: y * H, w: w * W, h: h * H };
      manualInner = val(manual, 'layoutTarget') === 'inner';
    }
  }
  rect(root, plot.x, plot.y, plot.w, plot.h, ownFill(child(plotArea, 'spPr'), env.colors, NOWHERE), parseLine(child(child(plotArea, 'spPr'), 'ln'), env.colors, NOWHERE), env, getDefs);

  if (isPie) {
    drawPie(root, groups[0], plot, base, locale, env, getDefs, palette);
    return;
  }

  const scatter = groups.some((g) => g.type === 'scatterChart');
  const horizontal = groups.some((g) => /bar/.test(g.type) && val(g.el, 'barDir') === 'bar');
  const catAxis = axes.find((a) => a.kind !== 'valAx') ?? (scatter ? axes.find((a) => a.kind === 'valAx' && (horizontal ? /[lr]/ : /[bt]/).test(a.pos)) : undefined);
  const valAxes = axes.filter((a) => a.kind === 'valAx' && a !== catAxis);
  const primary = valAxes[0];
  const secondary = valAxes.find((a) => a !== primary && groups.some((g) => g.axIds.includes(a.id)));

  // Values per axis, stacked where the group stacks.
  const scaleFor = new Map<string, Scale>();
  const axisValues: { axis: Axis; values: number[]; percent: boolean }[] = [];
  for (const axis of valAxes) {
    const used = groups.filter((g) => g.axIds.includes(axis.id) || (axis === primary && !valAxes.some((a) => g.axIds.includes(a.id))));
    const values: number[] = [0];
    let percent = false;
    for (const g of used) {
      const grouping = val(g.el, 'grouping') ?? 'clustered';
      if (/stacked/i.test(grouping) && grouping !== 'standard') {
        if (grouping === 'percentStacked') percent = true;
        const n = Math.max(...g.series.map((s) => s.values.length), 0);
        for (let i = 0; i < n; i++) {
          let pos = 0, neg = 0;
          for (const s of g.series) {
            const v = s.values[i] ?? 0;
            if (v >= 0) pos += v;
            else neg += v;
          }
          values.push(pos, neg);
        }
      } else {
        for (const s of g.series) for (const v of s.values) if (v !== null) values.push(v);
      }
    }
    if (values.length === 1) values.push(1);
    axisValues.push({ axis, values, percent });
  }
  const xValues = groups.flatMap((g) => g.series.flatMap((s) => (s.xValues ?? s.values.map((_, i) => i + 1)).filter((v): v is number => v !== null)));
  let scatterX = null as Scale | null;
  const computeScales = (valueLength: number, xLength: number) => {
    for (const { axis, values, percent } of axisValues) scaleFor.set(axis.id, valueScale(values, axis, percent, valueLength, axis.style.size));
    scatterX = scatter && catAxis ? valueScale(xValues, catAxis, false, xLength, catAxis.style.size * 3) : null;
  };
  // A first guess for the label widths; the scales are redone once the plot size is known.
  computeScales(Infinity, Infinity);

  const categories = allSeries.reduce<string[]>((best, s) => (s.categories.length > best.length ? s.categories : best), []);
  const count = Math.max(categories.length, ...allSeries.map((s) => s.values.length), 1);
  const fmtFor = (axis: Axis | undefined, series?: Series) => (axis && !axis.sourceLinked && axis.format ? axis.format : (series?.format ?? axis?.format ?? 'General'));

  // Room for axis labels when the layout is automatic.
  if (!manualInner) {
    const axisSpace = (axis: Axis | undefined, vertical: boolean, scale?: Scale) => {
      if (!axis || axis.deleted || val(axis.el, 'tickLblPos') === 'none') return 0;
      const titleSpace = axis.title ? axis.titleStyle.size * 1.8 : 0;
      if (vertical) {
        const labels = scale ? ticks(scale).map((v) => formatNumber(v, fmtFor(axis, allSeries[0]), locale)) : categories;
        return Math.max(0, ...labels.map((l) => measure(l, axis.style))) + 8 + titleSpace;
      }
      return axis.style.size * 2 + titleSpace;
    };
    const valueAxisVertical = !horizontal;
    const leftAxis = valueAxisVertical ? primary : catAxis;
    const bottomAxis = valueAxisVertical ? catAxis : primary;
    const leftSpace = axisSpace(leftAxis, true, valueAxisVertical ? scaleFor.get(primary?.id ?? '') : scatterX ?? undefined);
    const bottomSpace = axisSpace(bottomAxis, false);
    const rightSpace = secondary ? axisSpace(secondary, true, scaleFor.get(secondary.id)) : 0;
    plot = { x: plot.x + leftSpace, y: plot.y + 4, w: Math.max(10, plot.w - leftSpace - rightSpace - 4), h: Math.max(10, plot.h - bottomSpace - 4) };
  }
  computeScales(horizontal ? plot.w : plot.h, plot.w);

  const valueAt = (scale: Scale, axis: Axis | undefined, v: number) => {
    let t = (v - scale.min) / (scale.max - scale.min);
    if (axis?.reversed) t = 1 - t;
    return horizontal ? plot.x + t * plot.w : plot.y + plot.h - t * plot.h;
  };
  const midCat = groups.some((g) => /line|area/.test(g.type)) && primary?.crossBetween === 'midCat';
  const band = (horizontal ? plot.h : plot.w) / (midCat ? Math.max(1, count - 1) : count);
  const catAt = (i: number) => {
    const reversed = catAxis?.reversed ?? false;
    const k = reversed ? count - 1 - i : i;
    const offset = midCat ? k * band : (k + 0.5) * band;
    return horizontal ? plot.y + plot.h - offset : plot.x + offset;
  };

  // Gridlines and value labels.
  for (const axis of [primary, secondary]) {
    if (!axis) continue;
    const scale = scaleFor.get(axis.id)!;
    for (const v of ticks(scale)) {
      const p = valueAt(scale, axis, v);
      if (axis.grid && axis === primary) {
        if (horizontal) lineEl(root, p, plot.y, p, plot.y + plot.h, axis.grid, getDefs);
        else lineEl(root, plot.x, p, plot.x + plot.w, p, axis.grid, getDefs);
      }
      if (axis.deleted || val(axis.el, 'tickLblPos') === 'none') continue;
      const label = formatNumber(v, fmtFor(axis, allSeries[0]), locale);
      if (horizontal) textEl(root, label, p, plot.y + plot.h + axis.style.size * 0.9, axis.style, 'middle');
      else if (axis === secondary || axis.pos === 'r') textEl(root, label, plot.x + plot.w + 6, p, axis.style, 'start');
      else textEl(root, label, plot.x - 6, p, axis.style, 'end');
    }
    if (axis.title && !axis.deleted) {
      if (horizontal) textEl(root, axis.title, plot.x + plot.w / 2, plot.y + plot.h + axis.style.size * 2.2, axis.titleStyle, 'middle');
      else {
        const x = axis === secondary ? W - axis.titleStyle.size : Math.max(axis.titleStyle.size * 0.8, plot.x - (plot.x - left) + axis.titleStyle.size * 0.6);
        textEl(root, axis.title, x, plot.y + plot.h / 2, axis.titleStyle, 'middle', axis === secondary ? 90 : -90);
      }
    }
  }

  // Category axis line and labels (or the x value axis of scatter charts).
  if (catAxis && !catAxis.deleted) {
    const zero = primary ? valueAt(scaleFor.get(primary.id)!, primary, clamp(0, scaleFor.get(primary.id)!.min, scaleFor.get(primary.id)!.max)) : plot.y + plot.h;
    if (horizontal) lineEl(root, zero, plot.y, zero, plot.y + plot.h, catAxis.line, getDefs);
    else lineEl(root, plot.x, zero, plot.x + plot.w, zero, catAxis.line, getDefs);
    if (val(catAxis.el, 'tickLblPos') !== 'none') {
      if (scatterX) {
        for (const v of ticks(scatterX)) {
          const x = plot.x + ((v - scatterX.min) / (scatterX.max - scatterX.min)) * plot.w;
          textEl(root, formatNumber(v, catAxis.format, locale), x, plot.y + plot.h + catAxis.style.size * 0.9, catAxis.style, 'middle');
        }
      } else {
        const lowPos = val(catAxis.el, 'tickLblPos') === 'low';
        const skip = Math.max(1, Math.ceil((categories.length * catAxis.style.size * 2.2) / (horizontal ? plot.h * 2 : plot.w * 3)));
        categories.forEach((label, i) => {
          if (i % skip) return;
          const p = catAt(i);
          if (horizontal) textEl(root, label, plot.x - 6, p, catAxis.style, 'end');
          else wrapLabel(root, label, p, (lowPos ? plot.y + plot.h : zero) + catAxis.style.size * 1.1, band * skip - 4, catAxis.style);
        });
      }
    }
    if (catAxis.title) {
      if (horizontal) textEl(root, catAxis.title, Math.max(catAxis.titleStyle.size, plot.x - 40), plot.y + plot.h / 2, catAxis.titleStyle, 'middle', -90);
      else textEl(root, catAxis.title, plot.x + plot.w / 2, plot.y + plot.h + catAxis.style.size * 2 + catAxis.titleStyle.size, catAxis.titleStyle, 'middle');
    }
  }

  // Series.
  let seriesIndex = 0;
  const layer = svg('g');
  root.append(layer);
  const labelLayer = svg('g');
  for (const g of groups) {
    const axis = valAxes.find((a) => g.axIds.includes(a.id)) ?? primary;
    const scale = scaleFor.get(axis?.id ?? '') ?? { min: 0, max: 1, step: 0.1 };
    const grouping = val(g.el, 'grouping') ?? 'clustered';
    const stacked = /stacked/i.test(grouping) && grouping !== 'standard';
    const percent = grouping === 'percentStacked';
    const totals = Array.from({ length: count }, (_, i) => g.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0) || 1);
    const valueOf = (s: Series, i: number) => (s.values[i] === null || s.values[i] === undefined ? null : percent ? s.values[i]! / totals[i] : s.values[i]!);
    const firstIndex = seriesIndex;
    seriesIndex += g.series.length;
    const baseline = valueAt(scale, axis, clamp(0, scale.min, scale.max));

    if (/bar/.test(g.type)) {
      const gap = clamp((int(child(g.el, 'gapWidth'), 'val') ?? 150) / 100, 0, 5);
      const overlap = stacked ? 1 : clamp((int(child(g.el, 'overlap'), 'val') ?? 0) / 100, -1, 1);
      const n = stacked ? 1 : g.series.length;
      // The gap between clusters is measured in bar widths, like the overlap.
      const barW = band / (n - (n - 1) * overlap + gap || 1);
      const groupW = barW * (n - (n - 1) * overlap);
      const pos = new Array(count).fill(0), neg = new Array(count).fill(0);
      g.series.forEach((s, k) => {
        const color = s.fill ?? solid(palette(firstIndex + k));
        for (let i = 0; i < count; i++) {
          const v = valueOf(s, i);
          if (v === null) continue;
          let from = 0, to = v;
          if (stacked) {
            from = v >= 0 ? pos[i] : neg[i];
            to = from + v;
            if (v >= 0) pos[i] = to;
            else neg[i] = to;
          }
          const a = stacked ? valueAt(scale, axis, from) : baseline, b = valueAt(scale, axis, to);
          const center = catAt(i);
          const start = center - groupW / 2 + (stacked ? 0 : k * barW * (1 - overlap));
          const fill = s.points.get(i) ?? (val(g.el, 'varyColors') === '1' && g.series.length === 1 ? solid(palette(i)) : color);
          if (horizontal) rect(layer, Math.min(a, b), start, Math.abs(b - a), barW, fill, s.line, env, getDefs);
          else rect(layer, start, Math.min(a, b), barW, Math.abs(b - a), fill, s.line, env, getDefs);
          if (s.labels) {
            const text = labelText(s, i, v, s.labels, locale, totals[i]);
            const mid = start + barW / 2;
            const outside = !s.labels.pos || s.labels.pos === 'outEnd';
            const at = s.labels.pos === 'ctr' ? (a + b) / 2 : s.labels.pos === 'inBase' ? a : b;
            const dir = v >= 0 ? 1 : -1;
            if (horizontal) textEl(labelLayer, text, at + (outside ? 4 : -4) * dir, mid, s.labels.style, outside === v >= 0 ? 'start' : 'end');
            else textEl(labelLayer, text, mid, at - (outside ? s.labels.style.size * 0.7 : -s.labels.style.size * 0.7) * dir, s.labels.style, 'middle');
          }
        }
      });
    } else if (/line|area|scatter/.test(g.type)) {
      const isArea = /area/.test(g.type);
      const acc = new Array(count).fill(0);
      g.series.forEach((s, k) => {
        const color = palette(firstIndex + k);
        const points: [number, number][] = [];
        const from: [number, number][] = [];
        for (let i = 0; i < Math.max(count, s.values.length); i++) {
          let v = valueOf(s, i);
          if (v === null) continue;
          const lower = acc[i] ?? 0;
          if (stacked) {
            v += lower;
            acc[i] = v;
          }
          let x: number;
          if (scatterX && g.type === 'scatterChart') {
            const xv = s.xValues?.[i] ?? i + 1;
            x = plot.x + ((xv - scatterX.min) / (scatterX.max - scatterX.min)) * plot.w;
          } else x = catAt(i);
          points.push(horizontal ? [valueAt(scale, axis, v), x] : [x, valueAt(scale, axis, v)]);
          from.push(horizontal ? [valueAt(scale, axis, stacked ? lower : 0), x] : [x, valueAt(scale, axis, stacked ? lower : 0)]);
        }
        if (!points.length) return;
        if (isArea) {
          const back = [...from].reverse();
          const d = `M${points.map(([x, y]) => `${r2(x)} ${r2(y)}`).join('L')}L${back.map(([x, y]) => `${r2(x)} ${r2(y)}`).join('L')}Z`;
          const el = svg('path', { d });
          paintFill(el, s.fill ?? solid(color), getDefs, plot.w, plot.h, env);
          paintStroke(el, s.line, getDefs, plot.w, plot.h);
          layer.append(el);
        } else {
          const line = s.line ?? { fill: solid(color), width: 2.25 * PT, dash: 'solid', cap: 'round', join: 'round' };
          const scatterNoLine = g.type === 'scatterChart' && /marker$/i.test(val(g.el, 'scatterStyle') ?? '') && !s.line;
          if (!scatterNoLine && line.fill.type !== 'none') {
            const el = svg('path', { d: s.smooth ? smoothPath(points) : points.map(([x, y], i) => `${i ? 'L' : 'M'}${r2(x)} ${r2(y)}`).join(''), fill: 'none' });
            paintStroke(el, line, getDefs, plot.w, plot.h);
            layer.append(el);
          }
          const symbol = s.marker.symbol === 'auto' ? AUTO_MARKERS[(firstIndex + k) % AUTO_MARKERS.length] : s.marker.symbol;
          if (symbol !== 'none') {
            const lineColor = fillColor(line.fill) ?? color;
            for (const [x, y] of points) markerEl(layer, symbol, x, y, s.marker.size, s.marker.fill ?? solid(lineColor), s.marker.line ?? { ...line, width: 0.75 * PT, dash: 'solid' }, env, getDefs);
          }
        }
        if (s.labels) {
          points.forEach(([x, y], i) => {
            const v = valueOf(s, i);
            if (v === null) return;
            const posLabel = s.labels!.pos;
            const dy = posLabel === 'b' ? s.labels!.style.size : posLabel === 'ctr' ? 0 : -s.labels!.style.size;
            textEl(labelLayer, labelText(s, i, v, s.labels!, locale, totals[i]), x + (posLabel === 'r' ? 8 : 0), y + dy, s.labels!.style, posLabel === 'r' ? 'start' : 'middle');
          });
        }
      });
    }
  }
  root.append(labelLayer);
}

function ticks(scale: Scale): number[] {
  const out: number[] = [];
  const n = Math.min(200, Math.round((scale.max - scale.min) / scale.step));
  for (let i = 0; i <= n; i++) out.push(Math.round((scale.min + i * scale.step) * 1e9) / 1e9);
  return out;
}

function wrapLabel(parent: SVGElement, label: string, x: number, y: number, width: number, style: TextStyle): void {
  const words = label.split(/\s+/);
  const lines: string[] = [];
  let current = '';
  for (const word of words) {
    const next = current ? `${current} ${word}` : word;
    if (current && measure(next, style) > width) {
      lines.push(current);
      current = word;
    } else current = next;
  }
  if (current) lines.push(current);
  lines.slice(0, 3).forEach((line, i) => textEl(parent, line, x, y + i * style.size * 1.15, style, 'middle'));
}

function labelText(s: Series, i: number, v: number, labels: Labels, locale: string, total: number): string {
  const parts: string[] = [];
  if (labels.series) parts.push(s.name);
  if (labels.category) parts.push(s.categories[i] ?? '');
  if (labels.value) parts.push(formatNumber(s.values[i] ?? v, labels.format ?? s.format, locale));
  if (labels.percent) parts.push(formatNumber((s.values[i] ?? v) / (total || 1), '0%', locale));
  return parts.join('; ');
}

function drawPie(root: SVGSVGElement, g: Group, plot: { x: number; y: number; w: number; h: number }, base: TextStyle, locale: string, env: Env, defs: () => SVGDefsElement, palette: (i: number) => Color): void {
  const s = g.series[0];
  if (!s) return;
  const values = s.values.map((v) => Math.max(0, v ?? 0));
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const r = Math.max(4, Math.min(plot.w, plot.h) / 2 - 6);
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const hole = g.type === 'doughnutChart' ? clamp((int(child(g.el, 'holeSize'), 'val') ?? 50) / 100, 0.1, 0.9) * r : 0;
  let angle = ((int(child(g.el, 'firstSliceAng'), 'val') ?? 0) - 90) * (Math.PI / 180);
  const pt = (a: number, radius: number) => `${r2(cx + radius * Math.cos(a))} ${r2(cy + radius * Math.sin(a))}`;
  values.forEach((v, i) => {
    if (v <= 0) return;
    const sweep = (v / total) * Math.PI * 2;
    const end = angle + sweep;
    const large = sweep > Math.PI ? 1 : 0;
    let d: string;
    if (sweep >= Math.PI * 2 - 1e-6) {
      d = `M${pt(angle, r)}A${r2(r)} ${r2(r)} 0 1 1 ${pt(angle + Math.PI, r)}A${r2(r)} ${r2(r)} 0 1 1 ${pt(angle, r)}Z`;
      if (hole) d += `M${pt(angle, hole)}A${r2(hole)} ${r2(hole)} 0 1 0 ${pt(angle + Math.PI, hole)}A${r2(hole)} ${r2(hole)} 0 1 0 ${pt(angle, hole)}Z`;
    } else if (hole) {
      d = `M${pt(angle, r)}A${r2(r)} ${r2(r)} 0 ${large} 1 ${pt(end, r)}L${pt(end, hole)}A${r2(hole)} ${r2(hole)} 0 ${large} 0 ${pt(angle, hole)}Z`;
    } else {
      d = `M${r2(cx)} ${r2(cy)}L${pt(angle, r)}A${r2(r)} ${r2(r)} 0 ${large} 1 ${pt(end, r)}Z`;
    }
    const el = svg('path', { d, 'fill-rule': 'evenodd' });
    paintFill(el, s.points.get(i) ?? (val(g.el, 'varyColors') === '0' ? (s.fill ?? solid(palette(0))) : solid(palette(i))), defs, r * 2, r * 2, env);
    paintStroke(el, s.line ?? { fill: solid({ r: 255, g: 255, b: 255, a: 1 }), width: 1, dash: 'solid', cap: 'butt', join: 'round' }, defs, r * 2, r * 2);
    root.append(el);
    if (s.labels) {
      const mid = angle + sweep / 2;
      const at = s.labels.pos === 'outEnd' ? r + base.size : hole ? (r + hole) / 2 : r * 0.65;
      textEl(root, labelText(s, i, v, s.labels, locale, total), cx + at * Math.cos(mid), cy + at * Math.sin(mid), s.labels.style, 'middle');
    }
    angle = end;
  });
}



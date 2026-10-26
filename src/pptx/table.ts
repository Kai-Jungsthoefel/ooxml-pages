// SPDX-License-Identifier: MIT
// Tables (a:tbl in a graphic frame): grid, merged cells, cell fills and borders, and table styles with their
// regions (whole table, banded rows/columns, first/last row and column).
import { h } from '../dom.js';
import { colorIn, css } from './color.js';
import { type Env, type Source, px, r2 } from './env.js';
import { type Fill, type Line, cssFill, fillColor, ownFill, parseLine, refFill, refLine } from './fill.js';
import type { Box } from './shapes.js';
import { type Layer, type RunProps, renderParagraphs } from './text.js';
import { attr, child, children, clamp, int, parseXml } from './xml.js';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const ln = (w: number, color: string) => `<a:ln w="${w}"><a:solidFill>${color}</a:solidFill></a:ln>`;
const scheme = (name: string, mod = '') => `<a:schemeClr val="${name}">${mod}</a:schemeClr>`;
const border = (sides: string[], line: string) => `<a:tcBdr>${sides.map((s) => `<a:${s}>${line}</a:${s}>`).join('')}</a:tcBdr>`;

/**
 * The table styles PowerPoint uses most, for files whose tableStyles.xml does not spell them out
 * (PowerPoint only writes styles that are in use, and some generators write none).
 */
const BUILTIN: Record<string, string> = {
  // Medium Style 2 – Accent 1, PowerPoint's default.
  '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': `<a:tblStyle ${A}>
    <a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/>${scheme('dk1')}</a:tcTxStyle><a:tcStyle>${border(['left', 'right', 'top', 'bottom', 'insideH', 'insideV'], ln(12700, scheme('lt1')))}<a:fill><a:solidFill>${scheme('accent1', '<a:tint val="20000"/>')}</a:solidFill></a:fill></a:tcStyle></a:wholeTbl>
    <a:band1H><a:tcStyle><a:fill><a:solidFill>${scheme('accent1', '<a:tint val="40000"/>')}</a:solidFill></a:fill></a:tcStyle></a:band1H>
    <a:band1V><a:tcStyle><a:fill><a:solidFill>${scheme('accent1', '<a:tint val="40000"/>')}</a:solidFill></a:fill></a:tcStyle></a:band1V>
    <a:lastCol><a:tcTxStyle b="on">${scheme('lt1')}</a:tcTxStyle><a:tcStyle><a:fill><a:solidFill>${scheme('accent1')}</a:solidFill></a:fill></a:tcStyle></a:lastCol>
    <a:firstCol><a:tcTxStyle b="on">${scheme('lt1')}</a:tcTxStyle><a:tcStyle><a:fill><a:solidFill>${scheme('accent1')}</a:solidFill></a:fill></a:tcStyle></a:firstCol>
    <a:lastRow><a:tcTxStyle b="on">${scheme('lt1')}</a:tcTxStyle><a:tcStyle>${border(['top'], ln(38100, scheme('lt1')))}<a:fill><a:solidFill>${scheme('accent1')}</a:solidFill></a:fill></a:tcStyle></a:lastRow>
    <a:firstRow><a:tcTxStyle b="on">${scheme('lt1')}</a:tcTxStyle><a:tcStyle>${border(['bottom'], ln(38100, scheme('lt1')))}<a:fill><a:solidFill>${scheme('accent1')}</a:solidFill></a:fill></a:tcStyle></a:firstRow>
  </a:tblStyle>`,
  // No Style, Table Grid.
  '{5940675A-B579-460E-94D1-54222C63F5DA}': `<a:tblStyle ${A}><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/>${scheme('tx1')}</a:tcTxStyle><a:tcStyle>${border(['left', 'right', 'top', 'bottom', 'insideH', 'insideV'], ln(12700, scheme('tx1')))}</a:tcStyle></a:wholeTbl></a:tblStyle>`,
};

const builtinCache = new Map<string, Element | null>();

function tableStyle(id: string | null, env: Env): Element | null {
  if (!id) return null;
  const own = env.deck.tableStyles.get(id);
  if (own) return own;
  if (!builtinCache.has(id)) builtinCache.set(id, BUILTIN[id] ? (parseXml(BUILTIN[id])?.documentElement ?? null) : null);
  return builtinCache.get(id) ?? null;
}

interface Region {
  part: Element;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/** Style regions that cover a cell, lowest priority first. */
function regions(style: Element | null, flags: Record<string, boolean>, r: number, c: number, rows: number, cols: number): Region[] {
  if (!style) return [];
  const out: Region[] = [];
  const add = (name: string, r0: number, r1: number, c0: number, c1: number) => {
    const part = child(style, name);
    if (part) out.push({ part, r0, r1, c0, c1 });
  };
  add('wholeTbl', 0, rows - 1, 0, cols - 1);
  const firstBodyRow = flags.firstRow ? 1 : 0, lastBodyRow = flags.lastRow ? rows - 2 : rows - 1;
  const firstBodyCol = flags.firstCol ? 1 : 0, lastBodyCol = flags.lastCol ? cols - 2 : cols - 1;
  if (flags.bandCol && c >= firstBodyCol && c <= lastBodyCol) add((c - firstBodyCol) % 2 === 0 ? 'band1V' : 'band2V', 0, rows - 1, c, c);
  if (flags.bandRow && r >= firstBodyRow && r <= lastBodyRow) add((r - firstBodyRow) % 2 === 0 ? 'band1H' : 'band2H', r, r, 0, cols - 1);
  if (flags.lastCol && c === cols - 1) add('lastCol', 0, rows - 1, c, c);
  if (flags.firstCol && c === 0) add('firstCol', 0, rows - 1, 0, 0);
  if (flags.lastRow && r === rows - 1) add('lastRow', r, r, 0, cols - 1);
  if (flags.firstRow && r === 0) add('firstRow', 0, 0, 0, cols - 1);
  if (flags.lastRow && flags.firstCol && r === rows - 1 && c === 0) add('swCell', r, r, c, c);
  if (flags.lastRow && flags.lastCol && r === rows - 1 && c === cols - 1) add('seCell', r, r, c, c);
  if (flags.firstRow && flags.firstCol && r === 0 && c === 0) add('nwCell', r, r, c, c);
  if (flags.firstRow && flags.lastCol && r === 0 && c === cols - 1) add('neCell', r, r, c, c);
  return out;
}

type Side = 'left' | 'right' | 'top' | 'bottom';

function borderLine(el: Element | null, env: Env, source: Source): Line | undefined {
  if (!el) return undefined;
  const line = child(el, 'ln');
  if (line) return parseLine(line, env.colors, source) ?? { fill: { type: 'none' }, width: 0, dash: 'solid', cap: 'butt', join: 'round' };
  return refLine(child(el, 'lnRef'), env, source);
}

function cssBorder(line: Line | undefined): string {
  if (!line || line.fill.type === 'none') return 'none';
  const color = fillColor(line.fill);
  if (!color) return 'none';
  const style = /dot/i.test(line.dash) && !/dash/i.test(line.dash) ? 'dotted' : /dash/i.test(line.dash) ? 'dashed' : 'solid';
  return `${r2(Math.max(line.width, 0.5))}px ${style} ${css(color)}`;
}

function textLayer(part: Element, env: Env): Layer | null {
  const tx = child(part, 'tcTxStyle');
  if (!tx) return null;
  const run: RunProps = {};
  const b = attr(tx, 'b'), i = attr(tx, 'i');
  if (b) run.b = b === 'on';
  if (i) run.i = i === 'on';
  const fontRef = child(tx, 'fontRef');
  if (fontRef) {
    run.latin = attr(fontRef, 'idx') === 'major' ? '+mj-lt' : '+mn-lt';
    const color = colorIn(fontRef, env.colors);
    if (color) run.color = color;
  }
  const font = child(tx, 'font');
  if (font) run.latin = attr(child(font, 'latin'), 'typeface') ?? run.latin;
  const color = colorIn(tx, env.colors);
  if (color) run.color = color;
  return { run };
}

export function renderTable(tbl: Element, box: Box, env: Env, source: Source, base: Layer[]): HTMLElement {
  const tblPr = child(tbl, 'tblPr');
  const flags: Record<string, boolean> = {};
  for (const name of ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol']) flags[name] = attr(tblPr, name) === '1' || attr(tblPr, name) === 'true';
  const style = child(tblPr, 'tableStyle') ?? tableStyle(child(tblPr, 'tableStyleId')?.textContent?.trim() ?? null, env);

  const widths = children(child(tbl, 'tblGrid'), 'gridCol').slice(0, 200).map((c) => clamp(px(int(c, 'w') ?? 0), 0, 20000));
  const rows = children(tbl, 'tr').slice(0, 2000);
  const cols = widths.length;
  const table = h('table', { class: 'pptx-table' });
  table.style.left = `${r2(box.x)}px`;
  table.style.top = `${r2(box.y)}px`;
  table.style.width = `${r2(widths.reduce((a, b) => a + b, 0))}px`;
  if (box.rot) table.style.transform = `rotate(${r2(box.rot)}deg)`;
  const colgroup = h('colgroup', {}, ...widths.map((w) => h('col', { style: `width:${r2(w)}px` })));
  table.append(colgroup);
  const body = h('tbody');
  table.append(body);

  rows.forEach((tr, r) => {
    const row = h('tr');
    row.style.height = `${r2(clamp(px(int(tr, 'h') ?? 0), 0, 20000))}px`;
    let c = 0;
    for (const tc of children(tr, 'tc')) {
      const col = c++;
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') continue;
      if (col >= cols) break;
      const colSpan = clamp(int(tc, 'gridSpan') ?? 1, 1, cols - col);
      const rowSpan = clamp(int(tc, 'rowSpan') ?? 1, 1, rows.length - r);
      const tcPr = child(tc, 'tcPr');
      const covering = regions(style, flags, r, col, rows.length, cols);
      const cell = h('td', { colSpan, rowSpan });
      const s = cell.style;

      let fill: Fill | undefined = ownFill(tcPr, env.colors, source);
      if (!fill) {
        for (const region of [...covering].reverse()) {
          const tcStyle = child(region.part, 'tcStyle');
          const f = ownFill(child(tcStyle, 'fill'), env.colors, source) ?? refFill(child(tcStyle, 'fillRef'), env, source);
          if (f) {
            fill = f;
            break;
          }
        }
      }
      cssFill(cell, fill, env);

      const own: Record<Side, string> = { left: 'lnL', right: 'lnR', top: 'lnT', bottom: 'lnB' };
      const borderWidth: Record<Side, number> = { left: 0, right: 0, top: 0, bottom: 0 };
      for (const side of ['left', 'right', 'top', 'bottom'] as Side[]) {
        let line: Line | undefined = child(tcPr, own[side]) ? parseLine(child(tcPr, own[side]), env.colors, source) : undefined;
        if (!line) {
          for (const region of [...covering].reverse()) {
            const bdr = child(child(region.part, 'tcStyle'), 'tcBdr');
            const outer =
              side === 'left' ? col === region.c0 : side === 'right' ? col + colSpan - 1 >= region.c1 : side === 'top' ? r === region.r0 : r + rowSpan - 1 >= region.r1;
            const name = outer ? side : side === 'left' || side === 'right' ? 'insideV' : 'insideH';
            const found = borderLine(child(bdr, name), env, source);
            if (found) {
              line = found;
              break;
            }
          }
        }
        const value = cssBorder(line);
        if (value !== 'none') {
          s.setProperty(`border-${side}`, value);
          borderWidth[side] = Math.max(line!.width, 0.5);
        }
      }

      // PowerPoint draws borders centered on the grid lines without making cells bigger; a collapsed CSS
      // border adds half its width to the cell, so the padding gives that back.
      const pad = (name: string, fallback: number, side: Side) =>
        `${r2(Math.max(0, px(clamp(int(tcPr, name) ?? fallback, 0, 10_000_000)) - borderWidth[side] / 2))}px`;
      s.padding = `${pad('marT', 45720, 'top')} ${pad('marR', 91440, 'right')} ${pad('marB', 45720, 'bottom')} ${pad('marL', 91440, 'left')}`;
      const anchor = attr(tcPr, 'anchor');
      s.verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';

      const txBody = child(tc, 'txBody');
      if (txBody) {
        const layers = [...base];
        for (const region of covering) {
          const layer = textLayer(region.part, env);
          if (layer) layers.push(layer);
        }
        cell.append(renderParagraphs(txBody, layers, { env, source }).el);
      }
      row.append(cell);
    }
    body.append(row);
  });
  return table;
}

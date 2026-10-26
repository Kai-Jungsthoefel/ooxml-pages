// SPDX-License-Identifier: MIT
// Tables: column widths from the grid Word saved, cell margins, row heights, vertical merges, borders in the
// collapsed model (cell borders over table borders, the heavier one between neighbours) or, with cell spacing,
// the separate model (every cell keeps its own borders inside the table's), shading, vertical text, and rows
// that split across pages with repeated header rows.
import { type Flow, type GroupItem, type Item, type PlacedFloat, type RuleItem, type StoryLayout, layoutStory } from './flow.js';
import type { Block, Cell, Row, Table } from './model.js';
import { type Border, type CellMargins, type TableBorders, shadingColor, visible } from './props.js';
import { borderWidth } from './flow.js';
import { eighthPt, twip } from './xml.js';

const EPS = 0.01;
const DEFAULT_MARGIN = 108;

interface CellBox {
  cell: Cell;
  row: number;
  /** Rows this cell spans (vertical merge); 0 for continuation cells. */
  rowSpan: number;
  /** Grid columns. */
  col: number;
  span: number;
  /** Position relative to the table's left edge, width, content width. */
  x: number;
  w: number;
  mar: Required<CellMargins>;
  layout: StoryLayout | null;
  /** Height the content needs including margins. */
  need: number;
  vertical: number;
}

interface RowBox {
  row: Row;
  index: number;
  cells: CellBox[];
  height: number;
  /** Width of the border above the row (and below it for the last row). */
  borderTop: number;
  borderBottom: number;
  exact: boolean;
  header: boolean;
  cantSplit: boolean;
  /** Cell spacing in px (0: collapsed borders). Word stores half the gap between two cells. */
  spacing: number;
}

function margins(cell: Cell, row: Row, t: Table): Required<CellMargins> {
  const tm = row.ex?.cellMar ?? t.props.cellMar ?? {};
  const cm = cell.props.mar ?? {};
  return {
    top: twip(cm.top ?? tm.top ?? 0),
    bottom: twip(cm.bottom ?? tm.bottom ?? 0),
    left: twip(cm.left ?? tm.left ?? DEFAULT_MARGIN),
    right: twip(cm.right ?? tm.right ?? DEFAULT_MARGIN),
  };
}

/** Column widths in px for the space available. */
function columnWidths(t: Table, available: number): number[] {
  const count = Math.max(t.grid.length, ...t.rows.map((r) => r.cells.reduce((n, c) => Math.max(n, c.col + c.span), 0) + (r.props.gridAfter ?? 0)), 1);
  let widths = Array.from({ length: count }, (_, i) => twip(t.grid[i] ?? 0));
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum <= EPS || widths.some((w) => w <= 0)) {
    // No usable grid: take the cell widths of the rows, or share the space evenly.
    const fromCells = new Array<number>(count).fill(0);
    for (const row of t.rows) {
      for (const c of row.cells) {
        const w = c.props.width;
        if (w && w.type === 'dxa' && c.span === 1 && w.w > 0) fromCells[c.col] = Math.max(fromCells[c.col], twip(w.w));
        else if (w && w.type === 'pct' && c.span === 1 && w.w > 0) fromCells[c.col] = Math.max(fromCells[c.col], (available * w.w) / 5000);
      }
    }
    const known = fromCells.reduce((a, b) => a + b, 0);
    const missing = fromCells.filter((w) => w <= 0).length;
    const rest = Math.max(0, available - known);
    widths = fromCells.map((w) => (w > 0 ? w : missing ? rest / missing : available / count));
  }
  // A table width in percent follows the page; the grid was saved for some other width.
  const tw = t.props.width;
  const total = widths.reduce((a, b) => a + b, 0);
  let target: number | null = null;
  if (tw?.type === 'pct' && tw.w > 0) target = (available * tw.w) / 5000;
  if (target !== null && total > 0 && Math.abs(target - total) > 1) widths = widths.map((w) => (w * target!) / total);
  return widths;
}

function weight(b: Border | undefined): number {
  if (!visible(b)) return 0;
  const factor = /double|triple|thick/.test(b.style) ? 2 : 1;
  return b.size * factor;
}

/** A cell edge's border and whether the cell (or its table style region) sets it, rather than the table. */
interface EdgeBorder {
  b: Border | undefined;
  own: boolean;
}

/**
 * The border that wins between two neighbouring cells. Word draws the heavier of two borders the cells set
 * themselves, but a border only one of them sets loses against the table's inside border of the other
 * (measured against Word's own output).
 */
function resolve(a: EdgeBorder | null, b: EdgeBorder): EdgeBorder {
  if (!a) return b;
  if (a.own && b.own) return weight(b.b) > weight(a.b) ? b : a;
  if (a.own !== b.own) {
    const inherited = a.own ? b : a;
    const own = a.own ? a : b;
    return visible(inherited.b) ? inherited : own;
  }
  return weight(b.b) > weight(a.b) ? b : a;
}

function sameBorder(a: Border | undefined, b: Border | undefined): boolean {
  if (!visible(a) && !visible(b)) return true;
  if (!a || !b) return false;
  return a.style === b.style && a.size === b.size && a.color === b.color;
}

function translate(items: Item[], dx: number, dy: number): Item[] {
  return items.map((item) => {
    if (item.t === 'group') return { ...item, x: item.x + dx, y: item.y + dy };
    return { ...item, x: item.x + dx, y: item.y + dy } as Item;
  });
}

function translateFloats(floats: PlacedFloat[], dx: number, dy: number): PlacedFloat[] {
  return floats.map((f) => ({ ...f, x: f.x + dx, y: f.y + dy }));
}

/** Highest y at or above `limit` where a cell's content can be cut without cutting through a line. */
function cutPoint(items: Item[], limit: number): number {
  let cut = limit;
  for (let pass = 0; pass < 8; pass++) {
    let moved = false;
    for (const item of items) {
      const top = item.y;
      const bottom = item.t === 'line' ? item.y + item.line.h : item.t === 'group' ? item.y + item.h : item.y;
      if (item.t !== 'line' && item.t !== 'group') continue;
      if (top < cut - EPS && bottom > cut + EPS) {
        cut = top;
        moved = true;
      }
    }
    if (!moved) break;
  }
  return Math.max(0, cut);
}

/** Splits items at y: what is above stays, the rest moves up by `cut`. */
function splitItems(items: Item[], cut: number): [Item[], Item[]] {
  const head: Item[] = [], tail: Item[] = [];
  for (const item of items) {
    if (item.t === 'line' || item.t === 'group') {
      if (item.y < cut - EPS) head.push(item);
      else tail.push({ ...item, y: item.y - cut } as Item);
    } else if (item.t === 'fill') {
      if (item.y + item.h <= cut + EPS) head.push(item);
      else if (item.y >= cut - EPS) tail.push({ ...item, y: item.y - cut });
      else {
        head.push({ ...item, h: cut - item.y });
        tail.push({ ...item, y: 0, h: item.y + item.h - cut });
      }
    } else {
      const end = item.vertical ? item.y + item.len : item.y;
      if (end <= cut + EPS) head.push(item);
      else if (item.y >= cut - EPS) tail.push({ ...item, y: item.y - cut });
      else {
        head.push({ ...item, len: cut - item.y });
        tail.push({ ...item, y: 0, len: end - cut });
      }
    }
  }
  return [head, tail];
}

function spacingOf(w: { w: number; type: string } | undefined): number {
  return w && (w.type === 'dxa' || !w.type) && w.w > 0 ? twip(Math.min(w.w, 1440)) : 0;
}

export function layoutTable(flow: Flow, t: Table, _prev: Block | null): void {
  const available = flow.colW;
  const widths = columnWidths(t, available);
  const colX: number[] = [0];
  for (const w of widths) colX.push(colX[colX.length - 1] + w);
  const tableW = colX[colX.length - 1];

  // Horizontal position of the table's left edge relative to the column.
  const firstCell = t.rows[0]?.cells[0];
  const firstMargin = firstCell && t.rows[0] ? margins(firstCell, t.rows[0], t).left : twip(DEFAULT_MARGIN);
  const ind = t.props.ind && t.props.ind.type !== 'pct' ? twip(t.props.ind.w) : 0;
  const jc = t.rows[0]?.props.jc ?? t.props.jc ?? 'left';
  let left: number;
  if (jc === 'center') left = (available - tableW) / 2;
  else if (jc === 'right' || jc === 'end') left = available - tableW;
  else left = flow.ctx.doc.settings.compatMode >= 15 ? ind : ind - firstMargin;
  const floating = !!t.props.tblpPr;

  // Cells, with the cells of vertical merges joined to the cell where the merge starts.
  const rows: RowBox[] = t.rows.map((row, index) => ({
    row, index, cells: [], height: 0, borderTop: 0, borderBottom: 0, exact: row.props.hRule === 'exact', header: !!row.props.tblHeader, cantSplit: !!row.props.cantSplit,
    spacing: spacingOf(row.props.cellSpacing ?? t.props.cellSpacing),
  }));
  const openMerge = new Map<number, CellBox>();
  for (const rb of rows) {
    if (rb.row.props.hidden) continue;
    for (const cell of rb.row.cells) {
      const x = colX[Math.min(cell.col, widths.length)];
      const w = colX[Math.min(cell.col + cell.span, widths.length)] - x;
      const box: CellBox = { cell, row: rb.index, rowSpan: 1, col: cell.col, span: cell.span, x, w, mar: margins(cell, rb.row, t), layout: null, need: 0, vertical: 0 };
      const merge = cell.props.vMerge;
      if (merge === 'continue' || (merge && merge !== 'restart')) {
        const start = openMerge.get(cell.col);
        if (start && start.span === cell.span) {
          start.rowSpan++;
          box.rowSpan = 0;
        }
      } else if (merge === 'restart') {
        openMerge.set(cell.col, box);
      } else {
        openMerge.delete(cell.col);
      }
      if (merge !== 'restart' && box.rowSpan) for (let c = cell.col; c < cell.col + cell.span; c++) if (openMerge.get(c) !== box) openMerge.delete(c);
      rb.cells.push(box);
    }
  }

  // Cell contents.
  for (const rb of rows) {
    for (const box of rb.cells) {
      if (!box.rowSpan) continue;
      const dir = box.cell.props.textDirection;
      const contentW = Math.max(1, box.w - box.mar.left - box.mar.right);
      if (dir === 'btLr' || dir === 'tbRl' || dir === 'tbRlV' || dir === 'tbLrV') {
        const fixed = rb.row.props.height && rb.row.props.hRule !== 'auto' ? twip(rb.row.props.height) - box.mar.top - box.mar.bottom : 0;
        box.layout = layoutStory(flow.ctx, box.cell.blocks, fixed > 4 ? fixed : 2000);
        box.vertical = dir === 'btLr' ? -90 : 90;
        // Turned text needs the row as long as its longest line.
        box.need = (fixed > 4 ? fixed : box.layout.width) + box.mar.top + box.mar.bottom;
      } else {
        box.layout = layoutStory(flow.ctx, box.cell.blocks, contentW);
        box.need = box.layout.height + box.mar.top + box.mar.bottom;
      }
    }
  }

  // Borders of each cell's edges: own borders, then the table style's region borders, then the table's.
  const tableBorders = (rb: RowBox): TableBorders => rb.row.ex?.borders ?? t.props.borders ?? {};
  const cellAt = (r: number, col: number): CellBox | null => {
    const rb = rows[r];
    if (!rb) return null;
    for (const box of rb.cells) if (col >= box.col && col < box.col + box.span) return box;
    return null;
  };
  const owner = (box: CellBox | null): CellBox | null => {
    if (!box || box.rowSpan) return box;
    for (let r = box.row - 1; r >= 0; r--) {
      const above = cellAt(r, box.col);
      if (above && above.rowSpan) return above;
    }
    return box;
  };
  const edge = (box: CellBox, side: 'top' | 'bottom' | 'left' | 'right', rb: RowBox): EdgeBorder => {
    const own = box.cell.props.borders?.[side];
    if (own) return { b: own, own: true };
    const region = box.cell.edges?.[side];
    if (region) return { b: region, own: true };
    const tb = tableBorders(rb);
    const lastRow = rb.index + Math.max(1, box.rowSpan) - 1 >= rows.length - 1;
    switch (side) {
      case 'top':
        return { b: rb.index === 0 ? tb.top : tb.insideH, own: false };
      case 'bottom':
        return { b: lastRow ? tb.bottom : tb.insideH, own: false };
      case 'left':
        return { b: box.col === 0 ? tb.left : tb.insideV, own: false };
      case 'right':
        return { b: box.col + box.span >= widths.length ? tb.right : tb.insideV, own: false };
    }
  };

  /** Width of the horizontal border above a row (or below the last one): Word adds it to the row height. */
  const borderAbove = (r: number, below = false): number => {
    let w = 0;
    for (let g = 0; g < widths.length; g++) {
      const box = owner(cellAt(below ? r : r, g));
      if (!box) continue;
      let b: EdgeBorder;
      if (below) b = edge(box, 'bottom', rows[box.row]);
      else {
        const above = r > 0 ? owner(cellAt(r - 1, g)) : null;
        if (above === box) continue;
        b = resolve(above ? edge(above, 'bottom', rows[above.row]) : null, edge(box, 'top', rows[box.row]));
      }
      w = Math.max(w, borderWidth(b.b));
    }
    return w;
  };

  /** Separate model: the border a cell draws on one side (its own, else the table's inside or outer border). */
  const ownEdge = (box: CellBox, side: 'top' | 'bottom' | 'left' | 'right'): Border | undefined => {
    const own = box.cell.props.borders?.[side] ?? box.cell.edges?.[side];
    if (own) return own;
    const tb = tableBorders(rows[box.row]);
    return side === 'top' || side === 'bottom' ? (tb.insideH ?? tb[side]) : (tb.insideV ?? tb[side]);
  };
  /** Separate model: the space above a row's content (table border, spacing, cell border) and below it. */
  const separateTop = (rb: RowBox) =>
    (rb.index === 0 ? borderWidth(tableBorders(rb).top) : 0) + rb.spacing + Math.max(0, ...rb.cells.filter((c) => c.rowSpan).map((c) => borderWidth(ownEdge(c, 'top'))));
  const separateBottom = (rb: RowBox) =>
    (rb.index === rows.length - 1 ? borderWidth(tableBorders(rb).bottom) : 0) + rb.spacing + Math.max(0, ...rb.cells.filter((c) => c.rowSpan).map((c) => borderWidth(ownEdge(c, 'bottom'))));

  // Row heights: tallest single-row cell, the row's own height rule, then merged cells stretch their last row.
  for (const rb of rows) {
    rb.borderTop = rb.row.props.hidden ? 0 : rb.spacing ? separateTop(rb) : borderAbove(rb.index);
    rb.borderBottom = rb.spacing ? (rb.row.props.hidden ? 0 : separateBottom(rb)) : rb.index === rows.length - 1 ? borderAbove(rb.index, true) : 0;
    let h = 0;
    for (const box of rb.cells) if (box.rowSpan === 1) h = Math.max(h, box.need);
    const spec = rb.row.props.height ? twip(rb.row.props.height) : 0;
    if (rb.row.props.hidden) h = 0;
    else if (rb.exact && spec > 0) h = spec;
    else h = Math.max(h + rb.borderTop + rb.borderBottom, spec);
    rb.height = h;
  }
  for (const rb of rows) {
    for (const box of rb.cells) {
      if (box.rowSpan <= 1) continue;
      const span = rows.slice(rb.index, rb.index + box.rowSpan);
      const total = span.reduce((n, r) => n + r.height, 0);
      if (total + EPS < box.need) {
        const last = span[span.length - 1];
        if (!last.exact) last.height += box.need - total;
      }
    }
  }

  // Placement, row by row; rows that do not fit move to the next page or split there.
  const tableX = flow.colX + left;
  const headerRows = rows.filter((r, i) => r.header && rows.slice(0, i).every((p) => p.header));
  let firstOnPage = true;
  let pageTop = flow.y;
  let topRowOfPage = 0;

  const startPage = () => {
    pageTop = flow.y;
    firstOnPage = true;
  };

  /** Draws one row (or the part of it given by `slice`) at the current position. */
  const drawRow = (rb: RowBox, height: number, slices: Map<CellBox, { items: Item[]; floats: PlacedFloat[]; offset: number }> | null, isPageTop: boolean, isPageBottom: boolean) => {
    const y = flow.y;
    const background: Item[] = [];
    const content: Item[] = [];
    for (const box of rb.cells) {
      if (!box.rowSpan) continue;
      // Merged cells reach over the following rows on this page.
      let h = height;
      if (box.rowSpan > 1) {
        let r = rb.index + 1;
        let left = box.rowSpan - 1;
        while (left > 0 && r < rows.length) {
          h += rows[r].height;
          r++;
          left--;
        }
      }
      const shd = shadingColor(box.cell.props.shd) ?? shadingColor(t.props.shd);
      if (shd && rb.spacing) {
        const top = rb.borderTop - borderWidth(ownEdge(box, 'top')), bottom = rb.borderBottom - borderWidth(ownEdge(box, 'bottom'));
        background.push({ t: 'fill', x: tableX + box.x + rb.spacing, y: y + top, w: box.w - 2 * rb.spacing, h: h - top - bottom, color: shd });
      } else if (shd) background.push({ t: 'fill', x: tableX + box.x, y, w: box.w, h, color: shd });
      const slice = slices?.get(box);
      const layout = box.layout;
      if (!layout) continue;
      const items = slice ? slice.items : layout.items;
      const floats = slice ? slice.floats : layout.floats;
      const contentH = slice ? slice.offset : layout.height;
      const bt = rb.exact ? 0 : rb.borderTop;
      let dy = box.mar.top + bt;
      const valign = box.cell.props.vAlign;
      if (!slice && !box.vertical && (valign === 'center' || valign === 'bottom')) {
        const free = h - bt - box.mar.top - box.mar.bottom - contentH - (rb.spacing || rb.index === rows.length - 1 ? rb.borderBottom : 0);
        if (free > 0) dy += valign === 'center' ? free / 2 : free;
      }
      const cx = tableX + box.x + box.mar.left;
      if (box.vertical) {
        const boxW = box.w - box.mar.left - box.mar.right, boxH = h - box.mar.top - box.mar.bottom;
        // Turned text: lines run along the cell's height, anchored like the cell's vertical alignment says.
        const free = boxW - layout.height;
        const shift = valign === 'center' ? free / 2 : valign === 'bottom' ? free : 0;
        const group: GroupItem = { t: 'group', x: cx + (box.vertical === -90 ? shift : free - shift), y: y + box.mar.top, w: Math.min(boxW, layout.height), h: boxH, rotate: box.vertical, clip: false, items, floats };
        content.push(group);
        continue;
      }
      if (rb.exact && contentH > h - box.mar.top - box.mar.bottom + EPS) {
        content.push({ t: 'group', x: tableX + box.x, y, w: box.w, h, rotate: 0, clip: true, items: translate(items, box.mar.left, dy), floats: translateFloats(floats, box.mar.left, dy) });
      } else {
        content.push(...translate(items, cx, y + dy));
        for (const f of translateFloats(floats, cx, y + dy)) flow.floats.push(f);
      }
    }
    if (rb.spacing) {
      const rules: RuleItem[] = [];
      const rule = (x: number, yy: number, len: number, vertical: boolean, b: Border | undefined) => {
        if (visible(b) && len > EPS) rules.push({ t: 'rule', x, y: yy, len, vertical, b });
      };
      const isLast = rb.index === rows.length - 1;
      for (const box of rb.cells) {
        if (!box.rowSpan) continue;
        let h = height;
        if (box.rowSpan > 1) for (let r = rb.index + 1; r < rb.index + box.rowSpan && r < rows.length; r++) h += rows[r].height;
        const bt = ownEdge(box, 'top'), bb = ownEdge(box, 'bottom'), bl = ownEdge(box, 'left'), br = ownEdge(box, 'right');
        // Each border is drawn centred on its line, so lines sit half their width inside the cell's box.
        const top = y + rb.borderTop - borderWidth(bt) / 2;
        const lastRows = rb.index + Math.max(1, box.rowSpan) - 1;
        const bottom = y + h - (rows[lastRows]?.borderBottom ?? 0) + borderWidth(bb) / 2;
        const left = tableX + box.x + rb.spacing + borderWidth(bl) / 2, right = tableX + box.x + box.w - rb.spacing - borderWidth(br) / 2;
        if (!isPageTop || rb.index === box.row) rule(left - borderWidth(bl) / 2, top, right - left + borderWidth(bl) / 2 + borderWidth(br) / 2, false, bt);
        if (!isPageBottom) rule(left - borderWidth(bl) / 2, bottom, right - left + borderWidth(bl) / 2 + borderWidth(br) / 2, false, bb);
        rule(left, top, bottom - top, true, bl);
        rule(right, top, bottom - top, true, br);
      }
      // The table's own frame around the cells.
      const tb = tableBorders(rb);
      const frameTop = rb.index === 0 || isPageTop ? y + borderWidth(tb.top) / 2 : y;
      const frameBottom = isLast || isPageBottom ? y + height - borderWidth(tb.bottom) / 2 : y + height;
      if (rb.index === 0 || isPageTop) rule(tableX, frameTop, tableW, false, tb.top);
      if (isLast || isPageBottom) rule(tableX, frameBottom, tableW, false, tb.bottom);
      rule(tableX + borderWidth(tb.left) / 2, y, height, true, tb.left);
      rule(tableX + tableW - borderWidth(tb.right) / 2, y, height, true, tb.right);
      flow.items.push(...background, ...content, ...rules);
      flow.y += height;
      return;
    }
    // Borders: top edge per grid column, left edges per cell, and the right/bottom edges at the table's end.
    const rules: RuleItem[] = [];
    const hline = (x0: number, x1: number, yy: number, b: Border | undefined) => {
      if (!visible(b)) return;
      const last = rules[rules.length - 1];
      if (last && !last.vertical && Math.abs(last.y - yy) < EPS && Math.abs(last.x + last.len - x0) < 0.5 && sameBorder(last.b, b)) last.len = x1 - last.x;
      else rules.push({ t: 'rule', x: x0, y: yy, len: x1 - x0, vertical: false, b });
    };
    const vline = (xx: number, y0: number, y1: number, b: Border | undefined) => {
      if (!visible(b) || y1 - y0 <= EPS) return;
      rules.push({ t: 'rule', x: xx, y: y0, len: y1 - y0, vertical: true, b });
    };
    for (let g = 0; g < widths.length; g++) {
      const box = owner(cellAt(rb.index, g));
      if (!box) continue;
      if (box.row !== rb.index && !isPageTop) continue;
      const above = rb.index > 0 && !isPageTop ? owner(cellAt(rb.index - 1, g)) : null;
      let b = edge(box, 'top', rows[box.row]);
      if (above && above !== box) b = resolve(edge(above, 'bottom', rows[above.row]), b);
      else if (isPageTop && rb.index > 0 && !b.own) b = { b: tableBorders(rb).top ?? b.b, own: false };
      hline(tableX + colX[g], tableX + colX[g + 1], y, b.b);
    }
    for (const box of rb.cells) {
      const realOwner = owner(box);
      if (!realOwner) continue;
      let h = height;
      if (box.rowSpan > 1) for (let r = rb.index + 1; r < rb.index + box.rowSpan && r < rows.length; r++) h += rows[r].height;
      if (!box.rowSpan) continue;
      const leftNeighbour = box.col > 0 ? owner(cellAt(rb.index, box.col - 1)) : null;
      const lb = resolve(leftNeighbour ? edge(leftNeighbour, 'right', rows[leftNeighbour.row]) : null, edge(box, 'left', rb));
      vline(tableX + box.x, y, y + h, lb.b);
      if (box.col + box.span >= widths.length || !cellAt(rb.index, box.col + box.span)) vline(tableX + box.x + box.w, y, y + h, edge(box, 'right', rb).b);
    }
    // Continuation cells of a merge still need their left/right edges on rows that start a page.
    if (isPageTop) {
      for (const box of rb.cells) {
        if (box.rowSpan) continue;
        const o = owner(box);
        if (!o) continue;
        vline(tableX + o.x, y, y + height, edge(o, 'left', rows[o.row]).b);
        if (o.col + o.span >= widths.length) vline(tableX + o.x + o.w, y, y + height, edge(o, 'right', rows[o.row]).b);
      }
    }
    const isLast = rb.index === rows.length - 1;
    if (isLast || isPageBottom) {
      for (let g = 0; g < widths.length; g++) {
        const box = owner(cellAt(rb.index, g));
        if (!box) continue;
        const own = edge(box, 'bottom', rows[box.row]);
        hline(tableX + colX[g], tableX + colX[g + 1], y + height, isLast || own.own ? own.b : (tableBorders(rb).bottom ?? own.b));
      }
    }
    flow.items.push(...background, ...content, ...rules);
    flow.y += height;
  };

  const repeatHeaders = () => {
    for (const hr of headerRows) drawRow(hr, hr.height, null, hr === headerRows[0], false);
  };

  for (let i = 0; i < rows.length; i++) {
    const rb = rows[i];
    if (rb.row.props.hidden) continue;
    // Height this row needs on this page: merged cells starting here must fit with their rows.
    let need = rb.height;
    for (const box of rb.cells) if (box.rowSpan > 1) need = Math.max(need, rows.slice(i, i + box.rowSpan).reduce((n, r) => n + r.height, 0));
    const remaining = flow.bottom - flow.y;
    if (flow.paged && !floating && rb.height > remaining + EPS) {
      const canSplit = !rb.cantSplit && !rb.exact && !rb.header && rb.cells.every((c) => c.rowSpan <= 1) && remaining > 20;
      if (canSplit) {
        // Split the row: each cell keeps the lines that fit, the rest continues on the next page.
        let rest = rb;
        let restHeight = rb.height;
        let guard = 0;
        let parts = new Map<CellBox, { items: Item[]; floats: PlacedFloat[]; offset: number }>();
        for (const box of rb.cells) if (box.layout) parts.set(box, { items: box.layout.items, floats: box.layout.floats, offset: box.layout.height });
        while (restHeight > flow.bottom - flow.y + EPS && guard++ < 50) {
          const room = flow.bottom - flow.y;
          const head = new Map<CellBox, { items: Item[]; floats: PlacedFloat[]; offset: number }>();
          const tail = new Map<CellBox, { items: Item[]; floats: PlacedFloat[]; offset: number }>();
          let tailHeight = 0;
          let anyHead = false;
          for (const box of rest.cells) {
            const part = parts.get(box);
            if (!part) continue;
            const limit = room - box.mar.top - box.mar.bottom;
            const cut = cutPoint(part.items, Math.max(0, limit));
            const [a, b] = splitItems(part.items, cut);
            if (a.length) anyHead = true;
            head.set(box, { items: a, floats: part.floats, offset: cut });
            tail.set(box, { items: b, floats: [], offset: Math.max(0, part.offset - cut) });
            tailHeight = Math.max(tailHeight, part.offset - cut + box.mar.top + box.mar.bottom);
          }
          if (!anyHead && !firstOnPage) {
            flow.breakColumn();
            startPage();
            repeatHeaders();
            firstOnPage = false;
            continue;
          }
          drawRow(rest, Math.max(room, 0), head, firstOnPage && i > 0, true);
          flow.breakColumn();
          startPage();
          repeatHeaders();
          firstOnPage = headerRows.length === 0;
          parts = tail;
          restHeight = tailHeight;
          rest = { ...rest, height: tailHeight };
        }
        drawRow(rest, Math.max(restHeight, rb.exact ? rb.height : 0), parts, firstOnPage, false);
        firstOnPage = false;
        topRowOfPage = i;
        continue;
      }
      // Move to the next page unless the row already starts at the top of a column, where more room cannot be
      // had. This includes the first row of a table that begins near the bottom of a page.
      if (!flow.atColumnTop()) {
        flow.breakColumn();
        startPage();
        topRowOfPage = i;
        if (i > 0) {
          repeatHeaders();
          firstOnPage = headerRows.length === 0;
        }
      }
    }
    drawRow(rb, rb.height, null, (firstOnPage && i > 0) || (i === topRowOfPage && i > 0 && flow.y <= pageTop + EPS), false);
    firstOnPage = false;
    void need;
  }
}

export { eighthPt };

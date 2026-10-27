// SPDX-License-Identifier: MIT
// Writes Word samples that exercise the document renderer (test/fixtures/):
//   text-layout.docx          fonts, line spacing, indents, alignment, tabs, character formatting, lists, borders,
//                             columns, footnotes, endnotes, fields, headers and footers
//   sections-tables.docx      table styles, borders, merges, row heights, repeated header rows, nested tables,
//                             section breaks (landscape, columns, page numbers), first/even/odd headers
//   floating-pictures.docx    anchored pictures with every wrap type and anchor, rotation, flips, cropping,
//                             pictures in headers and footers and in table cells
//   forms-shapes-charts.docx  legacy form fields, check box content controls, symbol fonts, DrawingML shapes,
//                             text boxes, groups, a drawing canvas and charts
//   vml-legacy.docx           VML shapes and text boxes, VML pictures, a watermark, text frames and a drop cap
//   page-effects.docx         page color with a tiled picture, an art page border, Word 2010 text effects, a VML
//                             horizontal rule, a table with cell spacing at the foot of a page and a keep-with-next
//                             chain longer than a page
//   document.docx             the basics: headings, character formatting, a table, a picture and a page break
// Everything here is written from scratch, so the output can be published as CC0-1.0.
// Usage: node scripts/generate-docx-samples.mjs
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { crc32, deflateRawSync, deflateSync } from 'node:zlib';

const ROOT = path.resolve(import.meta.dirname, '..', 'test', 'fixtures');

// ---- Containers -------------------------------------------------------------------------------------------

/** ZIP with deflated entries and fixed timestamps, so regenerating gives identical files. */
function zip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of entries) {
    const plain = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const data = deflateRawSync(plain, { level: 9 });
    const nameBytes = Buffer.from(name, 'utf8');
    const crc = crc32(plain);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(plain.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    locals.push(local, nameBytes, data);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(plain.length, 24);
    central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

function png(w, h, fn) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(fn(x, y).map((v) => Math.max(0, Math.min(255, Math.round(v)))), y * (w * 4 + 1) + 1 + x * 4);
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}

// ---- Pictures (all drawn here) ----------------------------------------------------------------------------

const PICTURES = {
  // Sea, sky and sun; a frame line shows cropping and stretching.
  harbor: png(240, 150, (x, y) => {
    if (x < 2 || y < 2 || x > 237 || y > 147) return [40, 60, 90, 255];
    if ((x - 180) ** 2 + (y - 40) ** 2 < 22 ** 2) return [255, 210, 90, 255];
    const sea = y > 95 + 6 * Math.sin(x / 18);
    if (sea) return [30 + y / 4, 90 + y / 6, 150, 255];
    return [150 + y / 2, 190 + y / 4, 235, 255];
  }),
  // A round badge on transparent ground, for tight wrapping.
  badge: png(160, 160, (x, y) => {
    const d = Math.hypot(x - 79.5, y - 79.5);
    if (d > 78) return [0, 0, 0, 0];
    if (d > 70) return [200, 80, 40, 255];
    const stripe = Math.floor((x + y) / 20) % 2;
    return stripe ? [250, 220, 160, 255] : [240, 180, 90, 255];
  }),
  // Tall lighthouse-like shape.
  tower: png(80, 200, (x, y) => {
    const half = 12 + (y / 200) * 14;
    if (Math.abs(x - 40) < half && y > 30) return Math.floor(y / 30) % 2 ? [220, 60, 60, 255] : [245, 245, 245, 255];
    if (Math.abs(x - 40) < 14 && y > 12 && y <= 30) return [250, 230, 120, 255];
    return [205, 225, 245, 255];
  }),
  // Logo: two overlapping squares.
  logo: png(120, 60, (x, y) => {
    if (x > 10 && x < 50 && y > 10 && y < 50) return [47, 109, 181, 255];
    if (x > 35 && x < 75 && y > 15 && y < 55) return [226, 120, 58, 230];
    if (x > 82 && x < 115 && y > 25 && y < 35) return [60, 60, 60, 255];
    return [0, 0, 0, 0];
  }),
  // Pale texture to put behind text.
  paper: png(200, 120, (x, y) => [235 - (x % 20 < 2 ? 25 : 0), 242 - (y % 20 < 2 ? 25 : 0), 250, 255]),
  // Quadrants with numbers of dots, so cropping and flips are visible.
  quadrants: png(160, 120, (x, y) => {
    const q = (x >= 80 ? 1 : 0) + (y >= 60 ? 2 : 0);
    const colors = [[220, 70, 70], [70, 160, 80], [60, 110, 210], [230, 180, 40]];
    const [cx, cy] = [(x % 80) - 40, (y % 60) - 30];
    for (let i = 0; i <= q; i++) if ((cx - (i * 14 - q * 7)) ** 2 + cy ** 2 < 25) return [255, 255, 255, 255];
    return [...colors[q], 255];
  }),
  // Parchment-like tile for a page background.
  parchment: png(48, 48, (x, y) => {
    const n = Math.sin(x * 0.9 + y * 0.4) * 6 + Math.sin(y * 1.3 - x * 0.2) * 5;
    return [250 + n / 3, 238 + n, 205 + n, 255];
  }),
  // Small icon.
  icon: png(32, 32, (x, y) => ((x - 15.5) ** 2 + (y - 15.5) ** 2 < 15 ** 2 ? ((x - 15.5) ** 2 + (y - 15.5) ** 2 < 8 ** 2 ? [255, 255, 255, 255] : [62, 156, 92, 255]) : [0, 0, 0, 0])),
};

// ---- XML helpers ------------------------------------------------------------------------------------------

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};
const ROOT_NS = `xmlns:wpc="${NS.wpc}" xmlns:mc="${NS.mc}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:m="${NS.m}" xmlns:v="${NS.v}" xmlns:wp14="${NS.wp14}" xmlns:wp="${NS.wp}" xmlns:w10="${NS.w10}" xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:wpg="${NS.wpg}" xmlns:wps="${NS.wps}" mc:Ignorable="w14 wp14"`;
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';

const xml = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
/** Inches and points to twips and EMU. */
const IN = 1440;
const emuIn = (n) => Math.round(n * 914400);
const emuPt = (n) => Math.round(n * 12700);

function rels(list) {
  return xml(`<Relationships xmlns="${PKG}">${list.map(([id, type, target, external]) => `<Relationship Id="${id}" Type="${type.startsWith('http') ? type : `${RT}/${type}`}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`);
}

// ---- Runs and paragraphs (elements in schema order) ---------------------------------------------------------

/** Run properties: { style, font, fontEA, b, i, caps, smallCaps, strike, dstrike, vanish, color, themeColor, spacing, w, kern, position, sz, highlight, u, uColor, shd, vertAlign, lang }. */
function rPr(o = {}) {
  const parts = [
    o.style && `<w:rStyle w:val="${o.style}"/>`,
    o.font && `<w:rFonts w:ascii="${esc(o.font)}" w:hAnsi="${esc(o.font)}" w:cs="${esc(o.font)}"${o.fontEA ? ` w:eastAsia="${esc(o.fontEA)}"` : ''}/>`,
    o.b && '<w:b/>',
    o.i && '<w:i/>',
    o.caps && '<w:caps/>',
    o.smallCaps && '<w:smallCaps/>',
    o.strike && '<w:strike/>',
    o.dstrike && '<w:dstrike/>',
    o.vanish && '<w:vanish/>',
    o.color && `<w:color w:val="${o.color}"${o.themeColor ? ` w:themeColor="${o.themeColor}"` : ''}/>`,
    o.spacing !== undefined && `<w:spacing w:val="${o.spacing}"/>`,
    o.w && `<w:w w:val="${o.w}"/>`,
    o.kern && `<w:kern w:val="${o.kern}"/>`,
    o.position && `<w:position w:val="${o.position}"/>`,
    o.sz && `<w:sz w:val="${o.sz * 2}"/><w:szCs w:val="${o.sz * 2}"/>`,
    o.highlight && `<w:highlight w:val="${o.highlight}"/>`,
    o.u && `<w:u w:val="${o.u}"${o.uColor ? ` w:color="${o.uColor}"` : ''}/>`,
    o.shd && `<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`,
    o.vertAlign && `<w:vertAlign w:val="${o.vertAlign}"/>`,
    o.lang && `<w:lang w:val="${o.lang}"/>`,
  ].filter(Boolean);
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

/** A run: a string, or { t, ...rPr }, or raw XML { xml }. Tabs in text become w:tab. */
function run(r) {
  if (typeof r === 'string') r = { t: r };
  if (r.xml) return r.xml;
  const props = rPr(r);
  const body = String(r.t ?? '')
    .split('\t')
    .map((part) => (part ? `<w:t xml:space="preserve">${esc(part)}</w:t>` : ''))
    .join('<w:tab/>');
  return `<w:r>${props}${body}</w:r>`;
}

const tab = () => ({ xml: '<w:r><w:tab/></w:r>' });
const br = (type) => ({ xml: `<w:r><w:br${type ? ` w:type="${type}"` : ''}/></w:r>` });
const sym = (font, char, extra = {}) => ({ xml: `<w:r>${rPr(extra)}<w:sym w:font="${font}" w:char="${char}"/></w:r>` });

function border(name, b) {
  if (!b) return '';
  const [val, sz, color, space = 0] = b;
  return `<w:${name} w:val="${val}" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`;
}

/**
 * Paragraph properties: { style, keepNext, keepLines, pageBreakBefore, framePr, widow (false), num: [id, lvl],
 * border: { top, left, bottom, right, between }, shd, tabs: [[val, pos, leader]], before, after, line, rule,
 * ind: { left, right, firstLine, hanging }, contextual, jc, mark (rPr), sectPr }.
 */
function pPr(o = {}) {
  const ind = o.ind;
  const spacing = o.before !== undefined || o.after !== undefined || o.line !== undefined;
  const parts = [
    o.style && `<w:pStyle w:val="${o.style}"/>`,
    o.keepNext && '<w:keepNext/>',
    o.keepLines && '<w:keepLines/>',
    o.pageBreakBefore && '<w:pageBreakBefore/>',
    o.framePr && `<w:framePr ${o.framePr}/>`,
    o.widow === false && '<w:widowControl w:val="0"/>',
    o.num && `<w:numPr><w:ilvl w:val="${o.num[1] ?? 0}"/><w:numId w:val="${o.num[0]}"/></w:numPr>`,
    o.border && `<w:pBdr>${border('top', o.border.top)}${border('left', o.border.left)}${border('bottom', o.border.bottom)}${border('right', o.border.right)}${border('between', o.border.between)}</w:pBdr>`,
    o.shd && `<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`,
    o.tabs && `<w:tabs>${o.tabs.map(([val, pos, leader]) => `<w:tab w:val="${val}"${leader ? ` w:leader="${leader}"` : ''} w:pos="${pos}"/>`).join('')}</w:tabs>`,
    spacing && `<w:spacing${o.before !== undefined ? ` w:before="${o.before}"` : ''}${o.after !== undefined ? ` w:after="${o.after}"` : ''}${o.line !== undefined ? ` w:line="${o.line}" w:lineRule="${o.rule ?? 'auto'}"` : ''}/>`,
    ind && `<w:ind${ind.left !== undefined ? ` w:left="${ind.left}"` : ''}${ind.right !== undefined ? ` w:right="${ind.right}"` : ''}${ind.firstLine !== undefined ? ` w:firstLine="${ind.firstLine}"` : ''}${ind.hanging !== undefined ? ` w:hanging="${ind.hanging}"` : ''}/>`,
    o.contextual && '<w:contextualSpacing/>',
    o.jc && `<w:jc w:val="${o.jc}"/>`,
    o.mark && rPr(o.mark),
    o.sectPr,
  ].filter(Boolean);
  return parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : '';
}

/** Paragraph: string, or { runs, ...pPr }. */
function p(o) {
  if (typeof o === 'string') o = { runs: [o] };
  return `<w:p>${pPr(o)}${(o.runs ?? []).map(run).join('')}</w:p>`;
}

/** A complex field with a cached result: field('PAGE', '1', rPr). */
function field(instr, result, props = {}, extra = '') {
  const r = rPr(props);
  return {
    xml: `<w:r>${r}<w:fldChar w:fldCharType="begin">${extra}</w:fldChar></w:r><w:r>${r}<w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r><w:r>${r}<w:fldChar w:fldCharType="separate"/></w:r>${result === null ? '' : `<w:r>${r}<w:t xml:space="preserve">${esc(result)}</w:t></w:r>`}<w:r>${r}<w:fldChar w:fldCharType="end"/></w:r>`,
  };
}

function bookmark(id, name, inner) {
  return { xml: `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${inner.map(run).join('')}<w:bookmarkEnd w:id="${id}"/>` };
}

function hyperlink(rid, runs, anchor) {
  return { xml: `<w:hyperlink${rid ? ` r:id="${rid}"` : ''}${anchor ? ` w:anchor="${anchor}"` : ''} w:history="1">${runs.map(run).join('')}</w:hyperlink>` };
}

// ---- Tables -----------------------------------------------------------------------------------------------

/**
 * Table: { style, width: [w, type], jc, ind, borders: { top, left, bottom, right, insideH, insideV }, layout,
 * cellMar: { top, left, bottom, right }, look, grid: [twips…], rows }. Rows: { cells, height: [val, rule], header,
 * cantSplit }. Cells: string | { blocks|text, w, span, vMerge, shd, borders, vAlign, mar, dir, runs }.
 */
function table(t) {
  const tblBorders = t.borders && `<w:tblBorders>${border('top', t.borders.top)}${border('left', t.borders.left)}${border('bottom', t.borders.bottom)}${border('right', t.borders.right)}${border('insideH', t.borders.insideH)}${border('insideV', t.borders.insideV)}</w:tblBorders>`;
  const mar = (m, tag) => (m ? `<w:${tag}>${['top', 'left', 'bottom', 'right'].filter((k) => m[k] !== undefined).map((k) => `<w:${k} w:w="${m[k]}" w:type="dxa"/>`).join('')}</w:${tag}>` : '');
  const tblPr = [
    t.style && `<w:tblStyle w:val="${t.style}"/>`,
    t.tblpPr && `<w:tblpPr ${t.tblpPr}/>`,
    `<w:tblW w:w="${t.width?.[0] ?? 0}" w:type="${t.width?.[1] ?? 'auto'}"/>`,
    t.jc && `<w:jc w:val="${t.jc}"/>`,
    t.ind !== undefined && `<w:tblInd w:w="${t.ind}" w:type="dxa"/>`,
    tblBorders,
    t.layout && `<w:tblLayout w:type="${t.layout}"/>`,
    mar(t.cellMar, 'tblCellMar'),
    `<w:tblLook w:val="${t.look ?? '04A0'}" w:firstRow="${t.lookFirstRow ?? 1}" w:lastRow="${t.lookLastRow ?? 0}" w:firstColumn="${t.lookFirstCol ?? 1}" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`,
  ].filter(Boolean);
  const grid = `<w:tblGrid>${t.grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  const rows = t.rows
    .map((row) => {
      const trPr = [row.cantSplit && '<w:cantSplit/>', row.height && `<w:trHeight w:val="${row.height[0]}"${row.height[1] ? ` w:hRule="${row.height[1]}"` : ''}/>`, row.header && '<w:tblHeader/>'].filter(Boolean);
      let col = 0;
      const cells = row.cells
        .map((cell) => {
          if (typeof cell === 'string') cell = { text: cell };
          const span = cell.span ?? 1;
          const w = cell.w ?? t.grid.slice(col, col + span).reduce((a, b) => a + b, 0);
          col += span;
          const tcBorders = cell.borders && `<w:tcBorders>${border('top', cell.borders.top)}${border('left', cell.borders.left)}${border('bottom', cell.borders.bottom)}${border('right', cell.borders.right)}</w:tcBorders>`;
          const tcPr = [
            `<w:tcW w:w="${w}" w:type="dxa"/>`,
            span > 1 && `<w:gridSpan w:val="${span}"/>`,
            cell.vMerge && `<w:vMerge${cell.vMerge === 'restart' ? ' w:val="restart"' : ''}/>`,
            tcBorders,
            cell.shd && `<w:shd w:val="${cell.shdVal ?? 'clear'}" w:color="${cell.shdColor ?? 'auto'}" w:fill="${cell.shd}"/>`,
            mar(cell.mar, 'tcMar'),
            cell.dir && `<w:textDirection w:val="${cell.dir}"/>`,
            cell.vAlign && `<w:vAlign w:val="${cell.vAlign}"/>`,
          ].filter(Boolean);
          const content = cell.blocks ? cell.blocks.join('') : p(cell.para ?? { runs: cell.runs ?? [cell.text ?? ''], jc: cell.jc });
          return `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${content}</w:tc>`;
        })
        .join('');
      return `<w:tr>${trPr.length ? `<w:trPr>${trPr.join('')}</w:trPr>` : ''}${cells}</w:tr>`;
    })
    .join('');
  return `<w:tbl><w:tblPr>${tblPr.join('')}</w:tblPr>${grid}${rows}</w:tbl>`;
}

// ---- Sections ---------------------------------------------------------------------------------------------

/** Section properties: { headers: {default, first, even}, footers, type, w, h, landscape, margins, cols, pgNum, titlePg, vAlign, pgBorders, docGrid }. */
function sectPr(s = {}) {
  const refs = [
    ...Object.entries(s.headers ?? {}).map(([type, id]) => `<w:headerReference w:type="${type}" r:id="${id}"/>`),
    ...Object.entries(s.footers ?? {}).map(([type, id]) => `<w:footerReference w:type="${type}" r:id="${id}"/>`),
  ];
  const [w, h] = s.landscape ? [s.h ?? 15840, s.w ?? 12240] : [s.w ?? 12240, s.h ?? 15840];
  const m = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0, ...(s.margins ?? {}) };
  const cols = s.cols
    ? `<w:cols w:num="${s.cols.num}" w:space="${s.cols.space ?? 720}"${s.cols.sep ? ' w:sep="1"' : ''}${s.cols.list ? ' w:equalWidth="0"' : ''}>${(s.cols.list ?? []).map(([cw, sp]) => `<w:col w:w="${cw}"${sp !== undefined ? ` w:space="${sp}"` : ''}/>`).join('')}</w:cols>`
    : '<w:cols w:space="720"/>';
  return `<w:sectPr>${refs.join('')}${s.type ? `<w:type w:val="${s.type}"/>` : ''}<w:pgSz w:w="${w}" w:h="${h}"${s.landscape ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="${m.gutter}"/>${s.pgBorders ?? ''}${s.pgNum ? `<w:pgNumType${s.pgNum.fmt ? ` w:fmt="${s.pgNum.fmt}"` : ''}${s.pgNum.start !== undefined ? ` w:start="${s.pgNum.start}"` : ''}/>` : ''}${cols}${s.vAlign ? `<w:vAlign w:val="${s.vAlign}"/>` : ''}${s.titlePg ? '<w:titlePg/>' : ''}<w:docGrid w:linePitch="360"/></w:sectPr>`;
}

// ---- Drawings ---------------------------------------------------------------------------------------------

let docPrId = 1;

function picGraphic({ rid, w, h, name = 'Picture', crop, rot, flipH, flipV, line, geom = 'rect' }) {
  const src = crop ? `<a:srcRect l="${crop[0]}" t="${crop[1]}" r="${crop[2]}" b="${crop[3]}"/>` : '';
  const xfrmAttrs = `${rot ? ` rot="${Math.round(rot * 60000)}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
  return `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.pic}"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="0" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/>${src}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${line ?? ''}</pic:spPr></pic:pic></a:graphicData></a:graphic>`;
}

/** Layout space of a rotated box: the extra on each side (effectExtent). */
function rotatedExtent(w, h, rot) {
  if (!rot) return [0, 0];
  const a = (rot * Math.PI) / 180;
  const bw = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
  const bh = Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a));
  return [Math.max(0, Math.round((bw - w) / 2)), Math.max(0, Math.round((bh - h) / 2))];
}

function inline({ w, h, graphic, name = 'Picture', descr = '', rot }) {
  const [ex, ey] = rotatedExtent(w, h, rot);
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="${ex}" t="${ey}" r="${ex}" b="${ey}"/><wp:docPr id="${docPrId++}" name="${esc(name)}"${descr ? ` descr="${esc(descr)}"` : ''}/><wp:cNvGraphicFramePr/>${graphic}</wp:inline></w:drawing>`;
}

function position(axis, pos) {
  const tag = axis === 'h' ? 'positionH' : 'positionV';
  const inner = pos.align ? `<wp:align>${pos.align}</wp:align>` : `<wp:posOffset>${pos.offset ?? 0}</wp:posOffset>`;
  return `<wp:${tag} relativeFrom="${pos.rel}">${inner}</wp:${tag}>`;
}

/**
 * Floating drawing: { w, h, graphic, posH: { rel, align | offset }, posV, wrap: { type, side, polygon }, behind, z,
 * dist: [t, b, l, r] (EMU), rot }.
 */
function anchor({ w, h, graphic, posH, posV, wrap = { type: 'square' }, behind = false, z = 251659264, dist = [0, 0, 114300, 114300], name = 'Picture', rot, layoutInCell = true }) {
  const [ex, ey] = rotatedExtent(w, h, rot);
  let wrapXml;
  switch (wrap.type) {
    case 'none':
      wrapXml = '<wp:wrapNone/>';
      break;
    case 'topAndBottom':
      wrapXml = '<wp:wrapTopAndBottom/>';
      break;
    case 'tight':
    case 'through': {
      const poly = wrap.polygon ?? [[0, 0], [0, 21600], [21600, 21600], [21600, 0], [0, 0]];
      const tag = wrap.type === 'tight' ? 'wrapTight' : 'wrapThrough';
      wrapXml = `<wp:${tag} wrapText="${wrap.side ?? 'bothSides'}"><wp:wrapPolygon edited="0"><wp:start x="${poly[0][0]}" y="${poly[0][1]}"/>${poly.slice(1).map(([x, y]) => `<wp:lineTo x="${x}" y="${y}"/>`).join('')}</wp:wrapPolygon></wp:${tag}>`;
      break;
    }
    default:
      wrapXml = `<wp:wrapSquare wrapText="${wrap.side ?? 'bothSides'}"/>`;
  }
  return `<w:drawing><wp:anchor distT="${dist[0]}" distB="${dist[1]}" distL="${dist[2]}" distR="${dist[3]}" simplePos="0" relativeHeight="${z}" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="${layoutInCell ? 1 : 0}" allowOverlap="1"><wp:simplePos x="0" y="0"/>${position('h', posH)}${position('v', posV)}<wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="${ex}" t="${ey}" r="${ex}" b="${ey}"/>${wrapXml}<wp:docPr id="${docPrId++}" name="${esc(name)}"/><wp:cNvGraphicFramePr/>${graphic}</wp:anchor></w:drawing>`;
}

/** Wraps DrawingML that needs Word 2010 (shapes, groups, canvases) in a markup-compatibility block. */
const alt = (requires, drawing) => `<mc:AlternateContent><mc:Choice Requires="${requires}">${drawing}</mc:Choice></mc:AlternateContent>`;
const drawingRun = (drawing, requires) => ({ xml: `<w:r>${requires ? alt(requires, drawing) : drawing}</w:r>` });

const scheme = (name, mods = '') => `<a:schemeClr val="${name}">${mods}</a:schemeClr>`;
const rgb = (hex, mods = '') => `<a:srgbClr val="${hex}">${mods}</a:srgbClr>`;
const solid = (color) => `<a:solidFill>${color}</a:solidFill>`;
const shapeStyle = (accent = 'accent1') =>
  `<wps:style><a:lnRef idx="2">${scheme(accent, '<a:shade val="50000"/>')}</a:lnRef><a:fillRef idx="1">${scheme(accent)}</a:fillRef><a:effectRef idx="0">${scheme(accent)}</a:effectRef><a:fontRef idx="minor">${scheme('lt1')}</a:fontRef></wps:style>`;

/** A Word shape (wps:wsp): { w, h, prst, av, fill, line, style, text: [paragraph XML], anchor (t/ctr/b), x, y, rot, connector, flipH }. */
function wsp({ w, h, prst = 'rect', av = {}, fill = '', line = '', style = shapeStyle(), text, anchor: textAnchor = 't', x = 0, y = 0, rot, connector = false, flipH, flipV, txbxId, linked, name }) {
  const xfrmAttrs = `${rot ? ` rot="${Math.round(rot * 60000)}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
  const geom = `<a:prstGeom prst="${prst}"><a:avLst>${Object.entries(av).map(([k, v]) => `<a:gd name="${k}" fmla="val ${v}"/>`).join('')}</a:avLst></a:prstGeom>`;
  const nv = `${name ? `<wps:cNvPr id="${docPrId++}" name="${esc(name)}"/>` : ''}${connector ? '<wps:cNvCnPr/>' : `<wps:cNvSpPr${text || linked ? ' txBox="1"' : ''}/>`}`;
  const txbx = linked ? `<wps:linkedTxbx id="${linked[0]}" seq="${linked[1]}"/>` : text ? `<wps:txbx${txbxId !== undefined ? ` id="${txbxId}"` : ''}><w:txbxContent>${text.join('')}</w:txbxContent></wps:txbx>` : '';
  return `<wps:wsp>${nv}<wps:spPr><a:xfrm${xfrmAttrs}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>${geom}${fill}${line}</wps:spPr>${style}${txbx}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="${textAnchor}" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>`;
}

const shapeGraphic = (shape) => `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.wps}">${shape}</a:graphicData></a:graphic>`;
const groupGraphic = (w, h, children) =>
  `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.wpg}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/><a:chOff x="0" y="0"/><a:chExt cx="${w}" cy="${h}"/></a:xfrm></wpg:grpSpPr>${children.join('')}</wpg:wgp></a:graphicData></a:graphic>`;
const canvasGraphic = (children) =>
  `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.wpc}"><wpc:wpc><wpc:bg>${solid(rgb('F4F6F9'))}</wpc:bg><wpc:whole><a:ln w="9525">${solid(rgb('A6AEB8'))}</a:ln></wpc:whole>${children.join('')}</wpc:wpc></a:graphicData></a:graphic>`;
const chartGraphic = (rid) => `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="${rid}"/></a:graphicData></a:graphic>`;

/** A picture inside a group or canvas (pic:pic with its own offset). */
function groupPic({ rid, x, y, w, h }) {
  return `<pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${docPrId++}" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`;
}

// ---- Package parts ----------------------------------------------------------------------------------------

const FONT_LATIN = 'Calibri';
const FONT_MAJOR = 'Calibri Light';
const THEME_COLORS = [['dk1', null, 'windowText', '000000'], ['lt1', null, 'window', 'FFFFFF'], ['dk2', '44546A'], ['lt2', 'E7E6E6'], ['accent1', '4472C4'], ['accent2', 'ED7D31'], ['accent3', 'A5A5A5'], ['accent4', 'FFC000'], ['accent5', '5B9BD5'], ['accent6', '70AD47'], ['hlink', '0563C1'], ['folHlink', '954F72']];

function theme() {
  const colors = THEME_COLORS.map(([name, hex, sys, last]) => `<a:${name}>${sys ? `<a:sysClr val="${sys}" lastClr="${last}"/>` : `<a:srgbClr val="${hex}"/>`}</a:${name}>`).join('');
  const ph = (mods = '') => `<a:schemeClr val="phClr">${mods}</a:schemeClr>`;
  const grad = (a, b, c) =>
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${ph(a)}</a:gs><a:gs pos="50000">${ph(b)}</a:gs><a:gs pos="100000">${ph(c)}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
  return xml(
    `<a:theme xmlns:a="${NS.a}" name="Sample Theme"><a:themeElements><a:clrScheme name="Sample">${colors}</a:clrScheme><a:fontScheme name="Sample"><a:majorFont><a:latin typeface="${FONT_MAJOR}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${FONT_LATIN}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Sample"><a:fillStyleLst><a:solidFill>${ph()}</a:solidFill>${grad('<a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>', '<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/>', '<a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>')}${grad('<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>', '<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/>', '<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>')}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph()}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill>${ph()}</a:solidFill><a:solidFill>${ph('<a:tint val="95000"/><a:satMod val="170000"/>')}</a:solidFill>${grad('<a:tint val="93000"/>', '<a:tint val="98000"/>', '<a:shade val="63000"/>')}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
  );
}

function styles() {
  const style = (type, id, name, inner, extra = '') => `<w:style w:type="${type}"${extra} w:styleId="${id}"><w:name w:val="${name}"/>${inner}</w:style>`;
  const heading = (n, sz, color, before) =>
    style('paragraph', `Heading${n}`, `heading ${n}`, `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="0"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="${color}" w:themeColor="accent1" w:themeShade="${n === 3 ? '7F' : 'BF'}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`);
  const all = (sz, color) => ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => `<w:${k} w:val="single" w:sz="${sz}" w:space="0" w:color="${color}"/>`).join('');
  const grid4 =
    '<w:basedOn w:val="TableNormal"/><w:uiPriority w:val="49"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    `<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>${all(4, '8EAADB')}</w:tblBorders></w:tblPr>` +
    `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>` +
    '<w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="4472C4"/></w:tcBorders></w:tcPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr><w:tblStylePr w:type="lastCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="band1Vert"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>';
  return xml(
    `<w:styles xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      style('paragraph', 'Normal', 'Normal', '<w:qFormat/>', ' w:default="1"') +
      heading(1, 32, '2F5496', 240) +
      heading(2, 26, '2F5496', 40) +
      heading(3, 24, '1F3763', 40) +
      style('character', 'DefaultParagraphFont', 'Default Paragraph Font', '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>', ' w:default="1"') +
      style('table', 'TableNormal', 'Normal Table', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>', ' w:default="1"') +
      style('numbering', 'NoList', 'No List', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>', ' w:default="1"') +
      style('paragraph', 'Title', 'Title', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>') +
      style('paragraph', 'ListParagraph', 'List Paragraph', '<w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>') +
      style('paragraph', 'NoSpacing', 'No Spacing', '<w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>') +
      style('paragraph', 'Caption', 'caption', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="44546A" w:themeColor="text2"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>') +
      style('paragraph', 'Header', 'header', '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>') +
      style('paragraph', 'Footer', 'footer', '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>') +
      style('paragraph', 'FootnoteText', 'footnote text', '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>') +
      style('character', 'FootnoteReference', 'footnote reference', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>') +
      style('paragraph', 'EndnoteText', 'endnote text', '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>') +
      style('character', 'EndnoteReference', 'endnote reference', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>') +
      style('character', 'Hyperlink', 'Hyperlink', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr>') +
      style('character', 'PlaceholderText', 'Placeholder Text', '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="808080"/></w:rPr>') +
      style('paragraph', 'TOCHeading', 'TOC Heading', '<w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr>') +
      style('paragraph', 'TOC1', 'toc 1', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/></w:pPr>') +
      style('paragraph', 'TOC2', 'toc 2', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr>') +
      style('table', 'TableGrid', 'Table Grid', `<w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>${all(4, 'auto')}</w:tblBorders></w:tblPr>`) +
      style('table', 'GridTable4-Accent1', 'Grid Table 4 Accent 1', grid4) +
      '</w:styles>',
  );
}

function numbering() {
  const lvl = (ilvl, fmt, text, left, hanging, font, extra = '') =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${extra}<w:lvlText w:val="${esc(text)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
  const abstract = (id, levels) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels.join('')}</w:abstractNum>`;
  const nineLevels = (fn) => Array.from({ length: 9 }, (_, i) => fn(i));
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
  return xml(
    `<w:numbering xmlns:w="${NS.w}">` +
      // 0: 1. / 1.1. / 1.1.1.
      `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${nineLevels((i) => lvl(i, 'decimal', Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.', 360 + i * 432, 360 + i * 72))}</w:abstractNum>` +
      // 1: bullets • o ▪
      abstract(1, nineLevels((i) => lvl(i, 'bullet', ['', 'o', ''][i % 3], 720 + i * 720, 360, ['Symbol', 'Courier New', 'Wingdings'][i % 3]))) +
      // 2: a) i. 1.
      abstract(2, nineLevels((i) => lvl(i, ['lowerLetter', 'lowerRoman', 'decimal'][i % 3], ['%1)', `%${i + 1}.`, `(%${i + 1})`][i % 3].replace('%1', `%${i + 1}`), 720 + i * 720, 360))) +
      // 3: I. A. 1.
      abstract(3, nineLevels((i) => lvl(i, ['upperRoman', 'upperLetter', 'decimal'][i % 3], `%${i + 1}.`, 720 + i * 720, i % 3 === 0 ? 540 : 360))) +
      // 4: check mark bullets
      abstract(4, nineLevels((i) => lvl(i, 'bullet', '', 720 + i * 720, 360, 'Wingdings'))) +
      // 5: decimal with zero, ordinal text
      abstract(5, nineLevels((i) => lvl(i, i === 0 ? 'decimalZero' : fmts[i % 3], `${i === 0 ? 'Step ' : ''}%${i + 1}:`, 1080 + i * 720, 1080))) +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>' +
      '<w:num w:numId="4"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' +
      '<w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num><w:num w:numId="6"><w:abstractNumId w:val="4"/></w:num><w:num w:numId="7"><w:abstractNumId w:val="5"/></w:num>' +
      '</w:numbering>',
  );
}

function settings({ evenOdd = false, footnotes = false, background = false } = {}) {
  return xml(
    `<w:settings xmlns:w="${NS.w}" xmlns:m="${NS.m}"><w:zoom w:percent="100"/>${background ? '<w:displayBackgroundShape/>' : ''}<w:defaultTabStop w:val="720"/>${evenOdd ? '<w:evenAndOddHeaders/>' : ''}<w:characterSpacingControl w:val="doNotCompress"/>${footnotes ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr><w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>`,
  );
}

function notesPart(kind, notes) {
  const tag = kind === 'footnote' ? 'footnote' : 'endnote';
  const style = kind === 'footnote' ? 'Footnote' : 'Endnote';
  const sep = (id, type, el) => `<w:${tag} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${el}/></w:r></w:p></w:${tag}>`;
  return xml(
    `<w:${tag}s ${ROOT_NS}>${sep(-1, 'separator', 'separator')}${sep(0, 'continuationSeparator', 'continuationSeparator')}${notes
      .map(([id, text]) => `<w:${tag} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${style}Text"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${style}Reference"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${esc(text)}</w:t></w:r></w:p></w:${tag}>`)
      .join('')}</w:${tag}s>`,
  );
}

const noteRef = (kind, id) => ({ xml: `<w:r><w:rPr><w:rStyle w:val="${kind === 'footnote' ? 'Footnote' : 'Endnote'}Reference"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>` });

function headerPart(kind, blocks) {
  const tag = kind === 'header' ? 'hdr' : 'ftr';
  return xml(`<w:${tag} ${ROOT_NS}>${blocks.join('')}</w:${tag}>`);
}

function chartPart(kind) {
  const A = `xmlns:a="${NS.a}"`;
  const cats = ['North', 'East', 'South', 'West'];
  const series = [['2024', [42, 35, 28, 51]], ['2025', [48, 31, 36, 55]], ['2026', [53, 40, 33, 60]]];
  const strCache = (values) => `<c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (values) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const text = (sz, color = '595959', b = false) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"${b ? ' b="1"' : ''}>${solid(rgb(color))}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  const titleXml = (t) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0">${solid(rgb('404040'))}</a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  const ser = (i, [name, values], extra) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$1</c:f>${strCache([name])}</c:strRef></c:tx>${extra}<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f>${strCache(cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$2:$${String.fromCharCode(66 + i)}$5</c:f>${numCache(values)}</c:numRef></c:val></c:ser>`;
  const axes =
    `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525">${solid(rgb('BFBFBF'))}</a:ln></c:spPr>${text(900)}<c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525">${solid(rgb('D9D9D9'))}</a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${text(900)}<c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  let plot;
  let title;
  if (kind === 'bar') {
    title = 'Cargo by region';
    plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.map((s, i) => ser(i, s, `<c:spPr>${solid(scheme(`accent${i + 1}`))}</c:spPr><c:invertIfNegative val="0"/>`)).join('')}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="10"/><c:axId val="20"/></c:barChart>${axes}`;
  } else if (kind === 'line') {
    title = 'Arrivals per year';
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series
      .map((s, i) => ser(i, s, `<c:spPr><a:ln w="28575" cap="rnd">${solid(scheme(`accent${i + 1}`))}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${solid(scheme(`accent${i + 1}`))}<a:ln w="9525">${solid(scheme(`accent${i + 1}`))}</a:ln></c:spPr></c:marker>`))
      .join('')}<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>${axes}`;
  } else {
    title = 'Berth usage';
    const points = [0, 1, 2, 3].map((i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${solid(scheme(`accent${i + 1}`))}<a:ln w="19050">${solid(rgb('FFFFFF'))}</a:ln></c:spPr></c:dPt>`).join('');
    const labels = `<c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${text(1000, 'FFFFFF', true)}<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>`;
    plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, series[0], points + labels)}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  return xml(
    `<c:chartSpace xmlns:c="${NS.c}" ${A} xmlns:r="${NS.r}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${titleXml(title)}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:legend><c:legendPos val="${kind === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/>${text(900)}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr>${solid(rgb('FFFFFF'))}<a:ln w="9525">${solid(rgb('D9D9D9'))}</a:ln></c:spPr>${text(1000)}</c:chartSpace>`,
  );
}

const CREATED = '2026-10-26T09:00:00Z';

function docProps(title, subject) {
  return [
    ['docProps/core.xml', xml(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(title)}</dc:title><dc:subject>${esc(subject)}</dc:subject><dc:creator>ooxml-pages test generator</dc:creator><cp:lastModifiedBy>ooxml-pages test generator</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${CREATED}</dcterms:modified></cp:coreProperties>`)],
    ['docProps/app.xml', xml(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>ooxml-pages test generator</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>`)],
  ];
}

/**
 * A complete document. `parts`: { body, title, subject, headers: [[id, file, kind, blocks, rels]], notes: { footnote, endnote },
 * media: [[rid, name, data]], charts: [[rid, name, kind]], links: [[rid, url]], evenOdd, background (w:background XML) }.
 */
function docx({ body, title, subject, headers = [], notes = {}, media = [], charts = [], links = [], evenOdd = false, background = '' }) {
  const overrides = [
    ['/word/document.xml', `${CT}.wordprocessingml.document.main+xml`],
    ['/word/styles.xml', `${CT}.wordprocessingml.styles+xml`],
    ['/word/settings.xml', `${CT}.wordprocessingml.settings+xml`],
    ['/word/numbering.xml', `${CT}.wordprocessingml.numbering+xml`],
    ['/word/theme/theme1.xml', `${CT}.theme+xml`],
    ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['/docProps/app.xml', `${CT}.extended-properties+xml`],
  ];
  const docRels = [['rId1', 'styles', 'styles.xml'], ['rId2', 'settings', 'settings.xml'], ['rId3', 'numbering', 'numbering.xml'], ['rId4', 'theme', 'theme/theme1.xml']];
  const files = [];
  const hasNotes = !!(notes.footnote || notes.endnote);
  if (hasNotes) {
    docRels.push(['rId5', 'footnotes', 'footnotes.xml'], ['rId6', 'endnotes', 'endnotes.xml']);
    overrides.push(['/word/footnotes.xml', `${CT}.wordprocessingml.footnotes+xml`], ['/word/endnotes.xml', `${CT}.wordprocessingml.endnotes+xml`]);
    files.push(['word/footnotes.xml', notesPart('footnote', notes.footnote ?? [])], ['word/endnotes.xml', notesPart('endnote', notes.endnote ?? [])]);
  }
  for (const [id, file, kind, blocks, partRels = []] of headers) {
    docRels.push([id, kind, file]);
    overrides.push([`/word/${file}`, `${CT}.wordprocessingml.${kind}+xml`]);
    files.push([`word/${file}`, headerPart(kind, blocks)]);
    if (partRels.length) files.push([`word/_rels/${file}.rels`, rels(partRels.map(([rid, name]) => [rid, 'image', `media/${name}`]))]);
  }
  for (const [rid, name] of media) docRels.push([rid, 'image', `media/${name}`]);
  for (const [rid, name, kind] of charts) {
    docRels.push([rid, 'chart', `charts/${name}`]);
    overrides.push([`/word/charts/${name}`, `${CT}.drawingml.chart+xml`]);
    files.push([`word/charts/${name}`, chartPart(kind)]);
  }
  for (const [rid, url] of links) docRels.push([rid, 'hyperlink', url, true]);
  const usedMedia = new Map();
  for (const [, name, data] of media) usedMedia.set(name, data);
  for (const [, , , , partRels = []] of headers) for (const [, name, data] of partRels) usedMedia.set(name, data);
  const document = xml(`<w:document ${ROOT_NS}>${background}<w:body>${body.join('')}</w:body></w:document>`);
  return zip([
    ['[Content_Types].xml', xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')}</Types>`)],
    ['_rels/.rels', rels([['rId1', 'officeDocument', 'word/document.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', 'extended-properties', 'docProps/app.xml']])],
    ...docProps(title, subject),
    ['word/_rels/document.xml.rels', rels(docRels)],
    ['word/document.xml', document],
    ['word/styles.xml', styles()],
    ['word/settings.xml', settings({ evenOdd, footnotes: hasNotes, background: !!background })],
    ['word/numbering.xml', numbering()],
    ['word/theme/theme1.xml', theme()],
    ...files,
    ...[...usedMedia].map(([name, data]) => [`word/media/${name}`, data]),
  ]);
}

// ---- Text ------------------------------------------------------------------------------------------------

const S = [
  'The harbor master opened the logbook at dawn and noted the wind, the tide and the ships expected before noon.',
  'Every entry follows the same pattern: time, vessel, berth, cargo and the name of the pilot who brought her in.',
  'When fog rolls in from the estuary, the pilots wait at the outer buoy until the lighthouse signals that the channel is clear.',
  'Cranes along the eastern quay lift containers in a steady rhythm, while trucks queue at the gate with their engines idling.',
  'A small museum near the old warehouse shows charts drawn by hand, brass instruments and a model of the first steam tug.',
  "Visitors often ask why the tide tables change every year; the answer lies in the moon's long cycle and the shape of the bay.",
  'The repair yard keeps a list of spare parts, from propeller shafts to gaskets, so that no ship has to wait longer than necessary.',
  'In winter the ferry runs only twice a day, and passengers bring thermos flasks and stories about storms they barely remember.',
];
const text = (start, count) => Array.from({ length: count }, (_, i) => S[(start + i) % S.length]).join(' ');

const pageHeader = (label) =>
  headerPart('header', [p({ style: 'Header', runs: [{ t: label, color: '595959' }, tab(), tab(), { t: 'Page ', color: '595959' }, field('PAGE', '1', { color: '595959' }), { t: ' of ', color: '595959' }, field('NUMPAGES', '1', { color: '595959' })] })]);

// ---- Document 1: text layout --------------------------------------------------------------------------------

/** Office Math run with the math font, as Word writes it. */
const mr = (t, sty) => `<m:r>${sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : ''}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>${esc(t)}</m:t></m:r>`;

function textLayout() {
  const label = (t) => p({ runs: [{ t, b: true, sz: 9, color: '7F7F7F' }], after: 40, keepNext: true });
  const toc = [['Line spacing', '_Toc101', '1'], ['Indents and alignment', '_Toc102', '3'], ['Lists', '_Toc103', '4'], ['Columns', '_Toc104', '5']];
  const tocParas = toc.map(([title, name, page], i) =>
    p({
      style: 'TOC1',
      tabs: [['right', 9350, 'dot']],
      runs: [
        i === 0 && { xml: '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' },
        hyperlink(null, [title, tab(), field(`PAGEREF ${name} \\h`, page)], name),
        i === toc.length - 1 && { xml: '<w:r><w:fldChar w:fldCharType="end"/></w:r>' },
      ].filter(Boolean),
    }),
  );
  const heading = (title, name, extra = {}) => p({ style: 'Heading1', ...extra, runs: [bookmark(Number(name.slice(4)), name, [title])] });
  const quadratic =
    `<m:oMathPara><m:oMath>${mr('x=')}<m:f><m:num>${mr('-b±')}<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e>${mr('b')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup>${mr('-4ac')}</m:e></m:rad></m:num><m:den>${mr('2a')}</m:den></m:f></m:oMath></m:oMathPara>`;
  const area = `<m:oMath>${mr('A=π')}<m:sSup><m:e>${mr('r')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup></m:oMath>`;
  const body = [
    p({ style: 'Title', runs: ['Text layout test'] }),
    p({ runs: [{ t: 'Fonts, line spacing, indents, alignment, tabs, character formatting, lists, borders, columns, notes and fields.', i: true, color: '595959' }] }),
    ...tocParas,
    heading('Line spacing', '_Toc101'),
    label('Single (1.0)'),
    p({ runs: [text(0, 3)], line: 240 }),
    label('Multiple 1.15'),
    p({ runs: [text(1, 3)], line: 276 }),
    label('1.5 lines'),
    p({ runs: [text(2, 3)], line: 360 }),
    label('Double'),
    p({ runs: [text(3, 2)], line: 480 }),
    label('Exactly 12 pt'),
    p({ runs: [text(4, 2)], line: 240, rule: 'exact' }),
    label('At least 20 pt'),
    p({ runs: [text(5, 2)], line: 400, rule: 'atLeast' }),
    p({ style: 'Heading1', runs: ['Fonts'] }),
    ...[
      ['Arial', 11], ['Times New Roman', 12], ['Cambria', 11], ['Calibri Light', 14], ['Courier New', 10], ['Georgia', 11], ['Verdana', 9], ['Segoe UI', 10],
    ].map(([font, sz]) => p({ runs: [{ t: `${font} ${sz} pt: `, b: true, font, sz }, { t: S[0], font, sz }], after: 80 })),
    heading('Indents and alignment', '_Toc102', { pageBreakBefore: true }),
    label('Left and right indent 0.5 inch'),
    p({ runs: [text(6, 2)], ind: { left: 720, right: 720 } }),
    label('First line indent 0.5 inch'),
    p({ runs: [text(0, 2)], ind: { firstLine: 720 } }),
    label('Hanging indent 0.5 inch'),
    p({ runs: [text(1, 2)], ind: { left: 720, hanging: 720 } }),
    label('Centered'),
    p({ runs: [S[2]], jc: 'center' }),
    label('Right aligned'),
    p({ runs: [S[3]], jc: 'right' }),
    label('Justified'),
    p({ runs: [text(4, 3)], jc: 'both' }),
    label('Distributed'),
    p({ runs: [S[5]], jc: 'distribute' }),
    p({ style: 'Heading1', runs: ['Spacing'] }),
    p({ runs: ['Paragraph with 12 pt before and 18 pt after.'], before: 240, after: 360, shd: 'EEF3FA' }),
    p({ runs: ['Paragraph with 0 pt before and 6 pt after.'], before: 0, after: 120, shd: 'EEF3FA' }),
    p({ style: 'ListParagraph', runs: ['List paragraph one (contextual spacing, no space between)'] }),
    p({ style: 'ListParagraph', runs: ['List paragraph two'] }),
    p({ style: 'ListParagraph', runs: ['List paragraph three'] }),
    p({ style: 'Heading1', runs: ['Character formatting'] }),
    p({
      runs: [
        { t: 'Bold', b: true }, ', ', { t: 'italic', i: true }, ', ', { t: 'bold italic', b: true, i: true }, ', ', { t: 'underlined', u: 'single' }, ', ', { t: 'double underline', u: 'double' },
        ', ', { t: 'dotted underline', u: 'dotted' }, ', ', { t: 'wavy red underline', u: 'wave', uColor: 'C00000' }, ', ', { t: 'struck', strike: true }, ', ', { t: 'double struck', dstrike: true },
        ', ', { t: 'All Caps', caps: true }, ', ', { t: 'Small Caps', smallCaps: true }, ', H', { t: '2', vertAlign: 'subscript' }, 'O, E = mc', { t: '2', vertAlign: 'superscript' },
        ', ', { t: 'raised 3 pt', position: 6 }, ', ', { t: 'lowered 3 pt', position: -6 }, ', ', { t: 'expanded 2 pt', spacing: 40 }, ', ', { t: 'condensed 1 pt', spacing: -20 },
        ', ', { t: 'scaled 150 %', w: 150 }, ', ', { t: 'scaled 66 %', w: 66 }, ', ', { t: 'yellow highlight', highlight: 'yellow' }, ', ', { t: 'cyan highlight', highlight: 'cyan' },
        ', ', { t: 'shaded', shd: 'D9D9D9' }, ', ', { t: 'red text', color: 'C00000' }, ', ', { t: 'blue 14 pt', color: '2F5496', sz: 14 }, ', ', { t: 'hidden text', vanish: true }, 'and the end.',
      ],
    }),
    p({ style: 'Heading1', runs: ['Tabs'] }),
    p({ runs: ['Default tab stops:\tone\ttwo\tthree\tfour'] }),
    p({ tabs: [['left', 1440], ['center', 4320], ['right', 9360, 'dot']], runs: ['Left 1"\tleft\tcentered at 3"\tright at 6.5" with dots'] }),
    p({ tabs: [['decimal', 5040, 'hyphen'], ['right', 9360, 'underscore']], runs: ['Decimal 3.5":\t1,234.56\tunderscore'] }),
    p({ tabs: [['decimal', 5040, 'hyphen'], ['right', 9360, 'underscore']], runs: ['Decimal 3.5":\t7.5\tline'] }),
    p({ tabs: [['decimal', 5040, 'middleDot']], runs: ['Middle dots:\t42.125'] }),
    heading('Lists', '_Toc103', { pageBreakBefore: true }),
    p({ num: [1, 0], style: 'ListParagraph', runs: ['First level, decimal'] }),
    p({ num: [1, 1], style: 'ListParagraph', runs: ['Second level 1.1.'] }),
    p({ num: [1, 1], style: 'ListParagraph', runs: ['Second level 1.2.'] }),
    p({ num: [1, 2], style: 'ListParagraph', runs: ['Third level 1.2.1.'] }),
    p({ num: [1, 0], style: 'ListParagraph', runs: [`First level again. ${S[3]}`] }),
    p({ num: [2, 0], style: 'ListParagraph', runs: ['Bullet from Symbol'] }),
    p({ num: [2, 1], style: 'ListParagraph', runs: ['Bullet "o" in Courier New'] }),
    p({ num: [2, 2], style: 'ListParagraph', runs: ['Square bullet from Wingdings'] }),
    p({ num: [2, 0], style: 'ListParagraph', runs: [`Bullet with a long text: ${S[4]}`] }),
    p({ num: [3, 0], style: 'ListParagraph', runs: ['Letter a)'] }),
    p({ num: [3, 0], style: 'ListParagraph', runs: ['Letter b)'] }),
    p({ num: [3, 1], style: 'ListParagraph', runs: ['Roman ii. under b'] }),
    p({ runs: ['A paragraph between two lists.'] }),
    p({ num: [4, 0], style: 'ListParagraph', runs: ['Restarted list: a) again'] }),
    p({ num: [5, 0], style: 'ListParagraph', runs: ['Upper roman I.'] }),
    p({ num: [5, 0], style: 'ListParagraph', runs: ['Upper roman II.'] }),
    p({ num: [6, 0], style: 'ListParagraph', runs: ['Check mark bullet from Wingdings'] }),
    p({ num: [7, 0], runs: ['Numbered with leading zero and text before the number'] }),
    p({ num: [7, 0], runs: ['Second step'] }),
    p({ style: 'Heading1', runs: ['Borders and shading'] }),
    p({ border: { top: ['single', 8, '2F5496', 4], left: ['single', 8, '2F5496', 4], bottom: ['single', 8, '2F5496', 4], right: ['single', 8, '2F5496', 4], between: ['dotted', 4, '7F7F7F', 4] }, shd: 'F2F6FC', runs: [`Boxed paragraph one. ${S[1]}`] }),
    p({ border: { top: ['single', 8, '2F5496', 4], left: ['single', 8, '2F5496', 4], bottom: ['single', 8, '2F5496', 4], right: ['single', 8, '2F5496', 4], between: ['dotted', 4, '7F7F7F', 4] }, shd: 'F2F6FC', runs: ['Boxed paragraph two shares the box; a dotted line separates them.'] }),
    p({ border: { bottom: ['double', 6, 'C00000', 1] }, runs: ['Paragraph with a double red bottom border.'] }),
    p({ border: { left: ['thick', 24, 'ED7D31', 8] }, ind: { left: 360 }, runs: [`Thick orange bar on the left. ${S[6]}`] }),
    p({ style: 'Heading1', runs: ['Notes, links, fields and equations'] }),
    p({ runs: [{ xml: quadratic }] }),
    p({ runs: ['The area of a circle is ', { xml: area }, ', an equation inside a line of text.'] }),
    p({ runs: ['A sentence with a footnote', noteRef('footnote', 1), ', another one', noteRef('footnote', 2), ' and an endnote', noteRef('endnote', 1), '.'] }),
    p({ runs: ['External link: ', hyperlink('rId20', [{ t: 'example.com', style: 'Hyperlink' }]), '. Internal link: ', hyperlink(null, [{ t: 'jump to the columns section', style: 'Hyperlink' }], 'columns'), '.'] }),
    p({ runs: ['This is page ', field('PAGE', '3'), ' of ', field('NUMPAGES', '4'), '. Created: ', field('CREATEDATE \\@ "MMMM d, yyyy"', 'October 26, 2026'), '. Author: ', field('AUTHOR', 'ooxml-pages test generator'), '.'] }),
    p({ runs: [`Widow and orphan control keeps at least two lines of this paragraph together at a page break. ${text(2, 4)}`] }),
    p({ style: 'Heading1', pageBreakBefore: true, runs: [bookmark(1, 'columns', [bookmark(104, '_Toc104', ['Columns'])])] }),
    p({ runs: ['The next section has two columns with a line between them. It starts on the same page.'], sectPr: sectPr({ type: 'continuous', headers: { default: 'rId10' }, footers: { default: 'rId11' } }) }),
    p({ runs: [text(0, 4)] }),
    p({ runs: [text(4, 3)] }),
    p({ runs: [{ t: 'Before a column break.', b: true }, br('column')] }),
    p({ runs: [text(2, 3)] }),
    p({ runs: [text(5, 2)], sectPr: sectPr({ type: 'continuous', cols: { num: 2, space: 720, sep: true }, headers: { default: 'rId10' }, footers: { default: 'rId11' } }) }),
    p({ runs: ['Back to one column after a continuous section break.'] }),
    p({ runs: [text(1, 3)] }),
    sectPr({ type: 'continuous', headers: { default: 'rId10' }, footers: { default: 'rId11' } }),
  ];
  return docx({
    title: 'Text layout test',
    subject: 'Test document for the Word renderer: text layout',
    body,
    headers: [
      ['rId10', 'header1.xml', 'header', [p({ style: 'Header', runs: [{ t: 'Text layout test', color: '595959' }, tab(), tab(), { t: 'Page ', color: '595959' }, field('PAGE', '1', { color: '595959' }), { t: ' of ', color: '595959' }, field('NUMPAGES', '4', { color: '595959' })], border: { bottom: ['single', 4, 'A6A6A6', 1] } })]],
      ['rId11', 'footer1.xml', 'footer', [p({ style: 'Footer', jc: 'center', runs: [{ t: '– ', color: '7F7F7F' }, field('PAGE', '1', { color: '7F7F7F' }), { t: ' –', color: '7F7F7F' }] })]],
    ],
    notes: {
      footnote: [[1, 'The first footnote: a short remark at the bottom of the page.'], [2, `The second footnote is longer. ${S[7]}`]],
      endnote: [[1, 'An endnote appears at the end of the document.']],
    },
    links: [['rId20', 'https://example.com/']],
  });
}

// ---- Document 2: sections and tables ------------------------------------------------------------------------

function sectionsTables() {
  const b = (val = 'single', sz = 4, color = 'auto') => [val, sz, color, 0];
  const headerRow = ['Vessel', 'Berth', 'Arrival', 'Cargo (t)'];
  const ships = [
    ['Aurora', 'A1', '06:15', '1,250'], ['Borealis', 'A3', '07:40', '980'], ['Cygnus', 'B2', '09:05', '2,310'], ['Draco', 'C1', '10:30', '640'],
    ['Eridanus', 'B4', '11:55', '1,775'], ['Fornax', 'A2', '13:20', '1,020'], ['Grus', 'C3', '14:45', '3,150'], ['Hydra', 'B1', '16:10', '870'],
  ];
  const simple = table({
    style: 'TableGrid', width: [0, 'auto'], grid: [2340, 2340, 2340, 2340],
    rows: [{ cells: headerRow.map((t) => ({ runs: [{ t, b: true }], shd: 'D9D9D9' })) }, ...ships.slice(0, 4).map((r) => ({ cells: r.map((t, i) => ({ para: { runs: [t], jc: i === 3 ? 'right' : undefined } })) }))],
  });
  const styled = table({
    style: 'GridTable4-Accent1', width: [5000, 'pct'], grid: [2340, 2340, 2340, 2340], look: '04A0',
    rows: [{ cells: headerRow }, ...ships.slice(0, 6).map((r) => ({ cells: r })), { cells: ['Total', '', '', '7,975'] }],
  });
  const nested = table({
    style: 'TableGrid', grid: [800, 800], width: [1600, 'dxa'],
    rows: [{ cells: [{ text: 'in 1', shd: 'FFF2CC' }, 'in 2'] }, { cells: ['in 3', { text: 'in 4', shd: 'E2EFDA' }] }],
  });
  const merged = table({
    width: [9360, 'dxa'], grid: [1872, 1872, 1872, 1872, 1872], layout: 'fixed', jc: 'center',
    borders: { top: b('single', 12, '1F3763'), left: b('single', 12, '1F3763'), bottom: b('single', 12, '1F3763'), right: b('single', 12, '1F3763'), insideH: b('single', 4, '8496B0'), insideV: b('single', 4, '8496B0') },
    rows: [
      { cells: [{ text: 'Spans two columns', span: 2, shd: 'DEEAF6', jc: 'center' }, { text: 'Spans three columns, centered', span: 3, shd: 'FBE4D5', jc: 'center' }] },
      { height: [720, 'exact'], cells: [{ text: 'Merged down (restart), vertically centered', vMerge: 'restart', vAlign: 'center', shd: 'E2EFDA' }, { text: 'Exact row height 0.5"', vAlign: 'top' }, { text: 'Bottom', vAlign: 'bottom' }, { text: 'Center', vAlign: 'center' }, { blocks: [nested, p({ runs: ['After the nested table'] })] }] },
      { height: [1080, 'atLeast'], cells: [{ vMerge: 'continue' }, { text: 'Vertical text bottom to top', dir: 'btLr', shd: 'F2F2F2' }, { text: 'Dashed borders', borders: { top: b('dashed', 8, 'C00000'), left: b('dashed', 8, 'C00000'), bottom: b('dashed', 8, 'C00000'), right: b('dashed', 8, 'C00000') } }, { text: '25 % pattern', shd: 'FFFFFF', shdVal: 'pct25', shdColor: '2F5496' }, { text: 'Wide cell margins', mar: { top: 144, left: 288, bottom: 144, right: 288 } }] },
      { cells: [{ text: 'Double bottom border', span: 3, borders: { bottom: b('double', 6, '000000') } }, { text: 'No right border', span: 2, borders: { right: ['nil', 0, 'auto'] } }] },
    ],
  });
  const long = table({
    style: 'TableGrid', grid: [720, 2880, 2160, 3600], width: [9360, 'dxa'],
    rows: [
      { header: true, cells: ['#', 'Vessel', 'Berth', 'Remark'].map((t) => ({ runs: [{ t, b: true, color: 'FFFFFF' }], shd: '2F5496' })) },
      ...Array.from({ length: 34 }, (_, i) => ({
        cantSplit: i === 20,
        cells: [String(i + 1), ships[i % 8][0] + (i >= 8 ? ` ${Math.floor(i / 8) + 1}` : ''), ships[(i * 3) % 8][1], i === 20 ? `This row must not split across pages. ${S[2]} ${S[3]}` : i % 5 === 0 ? S[i % 8] : 'Scheduled'],
      })),
    ],
  });
  const floatingTable = table({
    style: 'TableGrid', grid: [1800, 1440], width: [3240, 'dxa'],
    tblpPr: 'w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="120"',
    rows: [{ cells: [{ runs: [{ t: 'Floating table', b: true }], shd: 'FFF2CC', span: 2 }] }, { cells: ['Tide high', '05:12'] }, { cells: ['Tide low', '11:27'] }],
  });
  const s1 = sectPr({ headers: { first: 'rId12', default: 'rId10', even: 'rId14' }, footers: { first: 'rId13', default: 'rId11', even: 'rId11' }, titlePg: true });
  const body = [
    p({ style: 'Title', runs: ['Sections and tables'] }),
    p({ runs: ['The first page has its own header and footer; odd and even pages differ after it.', { t: ' (Title page)', color: '7F7F7F' }] }),
    p({ style: 'Heading1', runs: ['Table Grid'] }),
    simple,
    p({ style: 'Caption', runs: ['Table 1: a simple grid with a shaded header row.'], before: 120 }),
    p({ style: 'Heading1', runs: ['Table style with banding'] }),
    styled,
    p({ style: 'Caption', runs: ['Table 2: header row, banded rows and a total row from the table style.'], before: 120 }),
    p({ style: 'Heading1', runs: ['Merged cells, borders, nesting'] }),
    merged,
    p({ runs: [''] }),
    floatingTable,
    p({ runs: [`Text next to a floating table. ${text(0, 3)}`] }),
    p({ style: 'Heading1', pageBreakBefore: true, runs: ['A long table with a repeated header row'] }),
    long,
    p({ runs: ['After the long table.'], sectPr: s1 }),
    p({ style: 'Heading1', runs: ['Landscape section'] }),
    p({ runs: [`This section is landscape with page numbers in lower roman, restarting at i, and a page border. ${text(3, 4)}`] }),
    p({ runs: [text(5, 5)], sectPr: sectPr({ landscape: true, pgNum: { fmt: 'lowerRoman', start: 1 }, headers: { default: 'rId15', even: 'rId15' }, footers: { default: 'rId16', even: 'rId16' }, pgBorders: '<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="2F5496"/><w:left w:val="single" w:sz="12" w:space="24" w:color="2F5496"/><w:bottom w:val="single" w:sz="12" w:space="24" w:color="2F5496"/><w:right w:val="single" w:sz="12" w:space="24" w:color="2F5496"/></w:pgBorders>' }) }),
    p({ style: 'Heading1', runs: ['Three columns'] }),
    p({ runs: [text(0, 6)] }),
    p({ runs: [text(3, 6)] }),
    p({ runs: [text(6, 4)], sectPr: sectPr({ cols: { num: 3, space: 360 }, pgNum: { start: 10 }, headers: { default: 'rId10', even: 'rId14' }, footers: { default: 'rId11', even: 'rId11' } }) }),
    p({ style: 'Heading1', runs: ['Vertically centered page'] }),
    p({ runs: ['This last section is centered vertically on its page and uses unequal columns.'] }),
    p({ runs: [text(1, 3)] }),
    sectPr({ type: 'nextPage', vAlign: 'center', cols: { num: 2, list: [[5760, 360], [3240]] }, headers: { default: 'rId10', even: 'rId14' }, footers: { default: 'rId11', even: 'rId11' } }),
  ];
  const hdr = (label, extra = []) => [p({ style: 'Header', runs: [{ t: label, b: true, color: '2F5496' }, tab(), tab(), 'Page ', field('PAGE', '1')], border: { bottom: ['single', 4, '2F5496', 1] } }), ...extra];
  return docx({
    title: 'Sections and tables',
    subject: 'Test document for the Word renderer: tables and sections',
    body,
    evenOdd: true,
    headers: [
      ['rId10', 'header1.xml', 'header', hdr('Odd page header')],
      ['rId11', 'footer1.xml', 'footer', [p({ style: 'Footer', jc: 'center', runs: ['Page ', field('PAGE', '2'), ' of ', field('NUMPAGES', '9'), ' · section pages ', field('SECTIONPAGES', '4')] })]],
      ['rId12', 'header2.xml', 'header', [p({ style: 'Header', jc: 'center', runs: [{ t: 'First page header', sz: 14, color: 'C00000' }] })]],
      ['rId13', 'footer2.xml', 'footer', [p({ style: 'Footer', jc: 'right', runs: [{ t: 'First page footer', i: true }] })]],
      ['rId14', 'header3.xml', 'header', hdr('Even page header')],
      ['rId15', 'header4.xml', 'header', [p({ style: 'Header', runs: [{ t: 'Landscape header', b: true }, tab(), tab(), 'Page ', field('PAGE', 'i')] })]],
      ['rId16', 'footer3.xml', 'footer', [p({ style: 'Footer', jc: 'center', runs: ['Landscape footer, page ', field('PAGE', 'i')] })]],
    ],
  });
}

// ---- Document 3: floating pictures --------------------------------------------------------------------------

function floatingPictures() {
  const pic = (rid, wIn, hIn, extra = {}) => ({ graphic: picGraphic({ rid, w: emuIn(wIn), h: emuIn(hIn), ...extra }), w: emuIn(wIn), h: emuIn(hIn) });
  const floatRun = (opts) => ({ xml: `<w:r>${anchor(opts)}</w:r>` });
  const inlineRun = (rid, wIn, hIn, extra = {}) => ({ xml: `<w:r>${inline({ ...pic(rid, wIn, hIn, extra), rot: extra.rot })}</w:r>` });
  const circle = Array.from({ length: 17 }, (_, i) => {
    const a = (i / 16) * Math.PI * 2;
    return [Math.round(10800 + 10600 * Math.cos(a)), Math.round(10800 + 10600 * Math.sin(a))];
  });
  const label = (t) => ({ t: `[${t}] `, b: true, color: '2F5496' });
  const body = [
    p({ style: 'Title', runs: ['Floating pictures'] }),
    p({ runs: [label('square, right of column'), floatRun({ ...pic('rId20', 1.6, 1.0), posH: { rel: 'column', align: 'right' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square' }, dist: [0, 0, emuIn(0.125), emuIn(0.125)] }), text(0, 5)] }),
    p({ runs: [label('square, left, 0.25" distance'), floatRun({ ...pic('rId22', 0.7, 1.75), posH: { rel: 'margin', align: 'left' }, posV: { rel: 'paragraph', offset: emuPt(6) }, wrap: { type: 'square' }, dist: [emuIn(0.25), emuIn(0.25), emuIn(0.25), emuIn(0.25)] }), text(2, 6)] }),
    p({ runs: [label('square, centered: text on both sides'), floatRun({ ...pic('rId25', 1.25, 0.94), posH: { rel: 'column', align: 'center' }, posV: { rel: 'paragraph', offset: emuPt(14) }, wrap: { type: 'square', side: 'bothSides' } }), text(4, 6)] }),
    p({ runs: [label('tight around a circle'), floatRun({ ...pic('rId21', 1.4, 1.4), posH: { rel: 'margin', offset: emuIn(2.2) }, posV: { rel: 'paragraph', offset: emuPt(4) }, wrap: { type: 'tight', polygon: circle } }), text(1, 7)] }),
    p({ style: 'Heading2', pageBreakBefore: true, runs: ['Wrap sides, through, top and bottom'] }),
    p({ runs: [label('wrap text only on the left side'), floatRun({ ...pic('rId20', 1.4, 0.875), posH: { rel: 'margin', offset: emuIn(2.5) }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square', side: 'left' } }), text(3, 4)] }),
    p({ runs: [label('wrap text only on the right side'), floatRun({ ...pic('rId20', 1.4, 0.875), posH: { rel: 'margin', offset: emuIn(2.5) }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square', side: 'right' } }), text(5, 4)] }),
    p({ runs: [label('largest side'), floatRun({ ...pic('rId25', 1.2, 0.9), posH: { rel: 'margin', offset: emuIn(1.0) }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square', side: 'largest' } }), text(7, 4)] }),
    p({ runs: [label('through'), floatRun({ ...pic('rId21', 1.0, 1.0), posH: { rel: 'column', align: 'left' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'through', polygon: circle } }), text(2, 3)] }),
    p({ runs: [label('top and bottom, centered on the margin'), floatRun({ ...pic('rId20', 2.4, 1.5), posH: { rel: 'margin', align: 'center' }, posV: { rel: 'paragraph', offset: emuPt(20) }, wrap: { type: 'topAndBottom' }, dist: [emuPt(6), emuPt(6), 0, 0] }), text(0, 3)] }),
    p({ style: 'Heading2', pageBreakBefore: true, runs: ['Behind and in front of text'] }),
    p({ runs: [label('behind text'), floatRun({ ...pic('rId23', 3.5, 1.4), posH: { rel: 'margin', offset: emuIn(1.5) }, posV: { rel: 'paragraph', offset: emuPt(8) }, wrap: { type: 'none' }, behind: true, z: 251650000 }), text(3, 5)] }),
    p({ runs: [label('in front of text, relative to the page, bottom right'), floatRun({ ...pic('rId21', 1.0, 1.0), posH: { rel: 'page', align: 'right' }, posV: { rel: 'page', align: 'bottom' }, wrap: { type: 'none' }, z: 251670000 }), text(6, 3)] }),
    p({ runs: [label('relative to the line, 0.1" above it, and the character'), floatRun({ ...pic('rId26', 0.3, 0.3), posH: { rel: 'character', offset: emuIn(0.1) }, posV: { rel: 'line', offset: emuIn(-0.1) }, wrap: { type: 'none' }, z: 251680000 }), text(4, 3)] }),
    p({ runs: [label('relative to the page: 5.9" from the left, 7.5" from the top, anchored here'), floatRun({ ...pic('rId20', 1.5, 0.94), posH: { rel: 'page', offset: emuIn(5.9) }, posV: { rel: 'page', offset: emuIn(7.5) }, wrap: { type: 'square' } }), text(1, 3)] }),
    p({ runs: [label('overlapping pictures, z order'), floatRun({ ...pic('rId25', 1.2, 0.9), posH: { rel: 'margin', offset: emuIn(0.3) }, posV: { rel: 'paragraph', offset: emuPt(30) }, wrap: { type: 'none' }, z: 251690000 }), floatRun({ ...pic('rId21', 0.9, 0.9), posH: { rel: 'margin', offset: emuIn(1.1) }, posV: { rel: 'paragraph', offset: emuPt(45) }, wrap: { type: 'none' }, z: 251691000 })] }),
    p({ runs: [''] }),
    p({ runs: [''] }),
    p({ runs: [''] }),
    p({ runs: [''] }),
    p({ style: 'Heading2', pageBreakBefore: true, runs: ['Rotation, flips, cropping, inline pictures'] }),
    p({ runs: [label('rotated 30°'), floatRun({ ...pic('rId25', 1.6, 1.2, { rot: 30 }), rot: 30, posH: { rel: 'margin', align: 'right' }, posV: { rel: 'paragraph', offset: emuPt(10) }, wrap: { type: 'square' } }), text(0, 4)] }),
    p({
      runs: [
        'Flipped horizontally ', inlineRun('rId25', 1.0, 0.75, { flipH: true }), ', vertically ', inlineRun('rId25', 1.0, 0.75, { flipV: true }), ', original ', inlineRun('rId25', 1.0, 0.75), '.',
      ],
    }),
    p({ runs: ['Cropped to the top right quadrant: ', inlineRun('rId25', 0.8, 0.6, { crop: [50000, 0, 0, 50000] }), ' and with a thick outline: ', inlineRun('rId25', 1.0, 0.75, { line: `<a:ln w="38100">${solid(rgb('C00000'))}</a:ln>` }), '.'] }),
    p({ runs: ['Inline picture rotated 90° in a line of text: ', inlineRun('rId22', 0.4, 1.0, { rot: 90 }), ' and an inline icon ', inlineRun('rId26', 0.2, 0.2), ' in the middle of a sentence.'] }),
    p({ runs: ['Ellipse-shaped picture: ', inlineRun('rId20', 1.4, 0.875, { geom: 'ellipse' }), ' and rounded: ', inlineRun('rId20', 1.4, 0.875, { geom: 'roundRect' })] }),
    p({ style: 'Heading2', runs: ['Picture in a table cell'] }),
    table({
      style: 'TableGrid', grid: [4680, 4680], width: [9360, 'dxa'],
      rows: [{ height: [1800, 'atLeast'], cells: [{ blocks: [p({ runs: [floatRun({ ...pic('rId25', 1.0, 0.75), posH: { rel: 'column', align: 'right' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square' } }), text(3, 2)] })] }, { blocks: [p({ runs: ['Inline in a cell: ', inlineRun('rId20', 1.6, 1.0)], jc: 'center' })] }] }],
    }),
    sectPr({ headers: { default: 'rId10' }, footers: { default: 'rId11' } }),
  ];
  const media = [['rId20', 'harbor.png', PICTURES.harbor], ['rId21', 'badge.png', PICTURES.badge], ['rId22', 'tower.png', PICTURES.tower], ['rId23', 'paper.png', PICTURES.paper], ['rId25', 'quadrants.png', PICTURES.quadrants], ['rId26', 'icon.png', PICTURES.icon]];
  const logo = { xml: `<w:r>${anchor({ graphic: picGraphic({ rid: 'rId1', w: emuIn(1.2), h: emuIn(0.6), name: 'Logo' }), w: emuIn(1.2), h: emuIn(0.6), posH: { rel: 'page', offset: emuIn(6.6) }, posV: { rel: 'page', offset: emuIn(0.3) }, wrap: { type: 'none' }, name: 'Logo' })}</w:r>` };
  const footerIcon = { xml: `<w:r>${inline({ graphic: picGraphic({ rid: 'rId1', w: emuIn(0.18), h: emuIn(0.18) }), w: emuIn(0.18), h: emuIn(0.18) })}</w:r>` };
  return docx({
    title: 'Floating pictures',
    subject: 'Test document for the Word renderer: anchored pictures',
    body,
    media,
    headers: [
      ['rId10', 'header1.xml', 'header', [p({ style: 'Header', runs: [logo, { t: 'Floating pictures – header with a logo anchored to the page', color: '595959' }] })], [['rId1', 'logo.png', PICTURES.logo]]],
      ['rId11', 'footer1.xml', 'footer', [p({ style: 'Footer', jc: 'center', runs: [footerIcon, ' Page ', field('PAGE', '1')] })], [['rId1', 'icon.png', PICTURES.icon]]],
    ],
  });
}

// ---- Document 4: forms, shapes, charts ----------------------------------------------------------------------

function formsShapesCharts() {
  let formId = 1;
  const checkbox = (checked, size) =>
    field('FORMCHECKBOX', null, {}, `<w:ffData><w:name w:val="Check${formId++}"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox>${size ? `<w:size w:val="${size}"/>` : '<w:sizeAuto/>'}<w:default w:val="${checked ? 1 : 0}"/></w:checkBox></w:ffData>`);
  const formText = (value) => field('FORMTEXT', value, {}, `<w:ffData><w:name w:val="Text${formId++}"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData>`);
  const dropdown = (entries, index) =>
    field('FORMDROPDOWN', null, {}, `<w:ffData><w:name w:val="Drop${formId++}"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="${index}"/>${entries.map((e) => `<w:listEntry w:val="${esc(e)}"/>`).join('')}</w:ddList></w:ffData>`);
  let sdtId = 100;
  const ccCheckbox = (checked, font = 'MS Gothic', on = '2612', off = '2610') => {
    const glyph = String.fromCodePoint(parseInt(checked ? on : off, 16));
    const ch = font === 'MS Gothic' ? glyph : String.fromCharCode(parseInt(checked ? on : off, 16));
    return {
      xml: `<w:sdt><w:sdtPr><w:id w:val="${sdtId++}"/><w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="${on}" w14:font="${font}"/><w14:uncheckedState w14:val="${off}" w14:font="${font}"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}"/></w:rPr><w:t>${esc(ch)}</w:t></w:r></w:sdtContent></w:sdt>`,
    };
  };
  const plainText = (value) => ({ xml: `<w:sdt><w:sdtPr><w:id w:val="${sdtId++}"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">${esc(value)}</w:t></w:r></w:sdtContent></w:sdt>` });
  const placeholder = (value) => ({ xml: `<w:sdt><w:sdtPr><w:id w:val="${sdtId++}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t xml:space="preserve">${esc(value)}</w:t></w:r></w:sdtContent></w:sdt>` });
  const formRow = (labelText, content) => ({ cells: [{ runs: [{ t: labelText, b: true }], shd: 'F2F2F2' }, { para: { runs: content } }] });
  const form = table({
    style: 'TableGrid', grid: [3000, 6360], width: [9360, 'dxa'],
    rows: [
      formRow('Name (legacy text field)', [formText('Jane Marlowe')]),
      formRow('Port (legacy drop-down)', [dropdown(['Hamburg', 'Rotterdam', 'Antwerp'], 1)]),
      formRow('Legacy check boxes', [checkbox(true), ' Crew list attached   ', checkbox(false), ' Hazardous cargo   ', checkbox(true, 28), ' 14 pt box']),
      formRow('Check box controls', [ccCheckbox(true), ' Customs cleared   ', ccCheckbox(false), ' Inspection due']),
      formRow('Check boxes (Wingdings)', [ccCheckbox(true, 'Wingdings', 'F0FE', 'F0A8'), ' Pilot ordered   ', ccCheckbox(false, 'Wingdings', 'F0FE', 'F0A8'), ' Tug ordered']),
      formRow('Text control', [plainText('Berth B4, north side')]),
      formRow('Empty control', [placeholder('Click or tap here to enter text.')]),
    ],
  });
  const symbols = [
    p({ runs: ['Wingdings: ', sym('Wingdings', 'F0FC'), ' check, ', sym('Wingdings', 'F0FB'), ' cross, ', sym('Wingdings', 'F0A8'), ' box, ', sym('Wingdings', 'F0FE'), ' checked box, ', sym('Wingdings', 'F06F'), ' square, ', sym('Wingdings', 'F0E0'), ' arrow, ', sym('Wingdings', 'F04A'), ' smiley, ', sym('Wingdings', 'F028'), ' phone.'] }),
    p({ runs: ['Wingdings 2: ', sym('Wingdings 2', 'F052'), ' checked, ', sym('Wingdings 2', 'F054'), ' crossed, ', sym('Wingdings 2', 'F0A3'), ' empty, ', sym('Wingdings 2', 'F050'), ' check mark, ', sym('Wingdings 2', 'F04F'), ' x mark.'] }),
    p({ runs: ['Symbol: ', sym('Symbol', 'F061'), sym('Symbol', 'F062'), sym('Symbol', 'F067'), ' ', sym('Symbol', 'F070'), ' ', sym('Symbol', 'F057'), ' ', sym('Symbol', 'F0B1'), ' ', sym('Symbol', 'F0A5'), ' ', sym('Symbol', 'F0B7'), ' ', sym('Symbol', 'F0E5'), ' ', sym('Symbol', 'F0D6'), ' (alpha beta gamma, pi, Omega, plus-minus, infinity, bullet, sum, root).'] }),
    p({ runs: ['Webdings: ', sym('Webdings', 'F061'), ' check, ', sym('Webdings', 'F072'), ' cross.'] }),
    p({ runs: ['Ballot boxes in Segoe UI Symbol: ', { t: '☐ ☑ ☒', font: 'Segoe UI Symbol' }, ' and in MS Gothic: ', { t: '☐ ☑ ☒', font: 'MS Gothic', fontEA: 'MS Gothic' }, ' and 16 pt: ', { t: '☐☒', font: 'Segoe UI Symbol', sz: 16 }] }),
    p({ runs: ['Wingdings text run: ', { t: 'üûþý¨o', font: 'Wingdings' }, ' (the same symbols typed in the Wingdings font).'] }),
  ];
  const tb = (paras) => paras.map((t) => p({ runs: [t], after: 60 }));
  const inlineShape = drawingRun(inline({ w: emuIn(2.2), h: emuIn(0.7), graphic: shapeGraphic(wsp({ w: emuIn(2.2), h: emuIn(0.7), prst: 'roundRect', text: [p({ runs: [{ t: 'Inline rounded rectangle', color: 'FFFFFF' }], jc: 'center', after: 0 })], anchor: 'ctr' })), name: 'Inline shape' }), 'wps');
  const floatingBox = drawingRun(
    anchor({
      w: emuIn(2.3), h: emuIn(1.5), posH: { rel: 'margin', align: 'right' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square' }, name: 'Text box',
      graphic: shapeGraphic(wsp({ w: emuIn(2.3), h: emuIn(1.5), fill: solid(rgb('FFF8E5')), line: `<a:ln w="12700">${solid(rgb('BF8F00'))}</a:ln>`, style: '', text: [p({ runs: [{ t: 'Text box', b: true }], after: 60 }), p({ num: [2, 0], style: 'ListParagraph', runs: ['bulleted line one'], after: 0 }), p({ num: [2, 0], style: 'ListParagraph', runs: ['bulleted line two'], after: 0 })] })),
    }),
    'wps',
  );
  const shapesRow = [
    ['ellipse', 'accent1', 'Ellipse'], ['rightArrow', 'accent2', 'Arrow'], ['star5', 'accent4', ''], ['wedgeRoundRectCallout', 'accent6', 'Callout'], ['flowChartDecision', 'accent5', 'Yes?'],
  ].map(([prst, accent, label], i) =>
    drawingRun(
      anchor({ w: emuIn(1.1), h: emuIn(0.8), posH: { rel: 'margin', offset: emuIn(i * 1.3) }, posV: { rel: 'paragraph', offset: emuPt(6) }, wrap: { type: 'topAndBottom' }, name: `Shape ${prst}`, z: 251700000 + i, graphic: shapeGraphic(wsp({ w: emuIn(1.1), h: emuIn(0.8), prst, style: shapeStyle(accent), text: label ? [p({ runs: [{ t: label, color: 'FFFFFF', sz: 10 }], jc: 'center', after: 0 })] : undefined, anchor: 'ctr' })) }),
      'wps',
    ),
  );
  const line = drawingRun(
    anchor({ w: emuIn(2.5), h: emuIn(0.5), posH: { rel: 'margin', offset: emuIn(0.2) }, posV: { rel: 'paragraph', offset: emuPt(4) }, wrap: { type: 'topAndBottom' }, name: 'Arrow line', graphic: shapeGraphic(wsp({ w: emuIn(2.5), h: emuIn(0.5), prst: 'line', connector: true, style: '', line: `<a:ln w="19050">${solid(rgb('C00000'))}<a:headEnd type="oval"/><a:tailEnd type="triangle" w="med" len="med"/></a:ln>` })) }),
    'wps',
  );
  const rotated = drawingRun(
    anchor({ w: emuIn(1.6), h: emuIn(0.6), rot: -15, posH: { rel: 'margin', offset: emuIn(3.6) }, posV: { rel: 'paragraph', offset: emuPt(4) }, wrap: { type: 'none' }, name: 'Rotated', graphic: shapeGraphic(wsp({ w: emuIn(1.6), h: emuIn(0.6), rot: -15, prst: 'rect', style: shapeStyle('accent2'), text: [p({ runs: [{ t: 'Rotated −15°', color: 'FFFFFF' }], jc: 'center', after: 0 })], anchor: 'ctr' })) }),
    'wps',
  );
  const groupW = emuIn(3.0), groupH = emuIn(1.4);
  const group = drawingRun(
    anchor({
      w: groupW, h: groupH, posH: { rel: 'margin', align: 'left' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square' }, name: 'Group',
      graphic: groupGraphic(groupW, groupH, [
        wsp({ x: 0, y: 0, w: emuIn(1.4), h: emuIn(1.4), prst: 'ellipse', style: shapeStyle('accent1'), name: 'Group ellipse' }),
        wsp({ x: emuIn(1.0), y: emuIn(0.35), w: emuIn(2.0), h: emuIn(0.7), prst: 'rect', fill: solid(rgb('FFFFFF')), line: `<a:ln w="12700">${solid(scheme('accent1'))}</a:ln>`, style: '', text: [p({ runs: ['Grouped text box'], jc: 'center', after: 0 })], anchor: 'ctr', name: 'Group text' }),
        groupPic({ rid: 'rId21', x: emuIn(0.15), y: emuIn(0.15), w: emuIn(0.5), h: emuIn(0.5) }),
      ]),
    }),
    'wpg',
  );
  const canvasW = emuIn(4.0), canvasH = emuIn(1.3);
  const canvas = drawingRun(
    inline({
      w: canvasW, h: canvasH, name: 'Canvas',
      graphic: canvasGraphic([
        wsp({ x: emuIn(0.2), y: emuIn(0.3), w: emuIn(1.2), h: emuIn(0.7), prst: 'roundRect', style: shapeStyle('accent3'), text: [p({ runs: [{ t: 'Start', color: 'FFFFFF' }], jc: 'center', after: 0 })], anchor: 'ctr', name: 'Canvas start' }),
        wsp({ x: emuIn(2.6), y: emuIn(0.3), w: emuIn(1.2), h: emuIn(0.7), prst: 'roundRect', style: shapeStyle('accent6'), text: [p({ runs: [{ t: 'Finish', color: 'FFFFFF' }], jc: 'center', after: 0 })], anchor: 'ctr', name: 'Canvas finish' }),
        wsp({ x: emuIn(1.4), y: emuIn(0.65), w: emuIn(1.2), h: 0, prst: 'straightConnector1', connector: true, style: '', line: `<a:ln w="15875">${solid(rgb('404040'))}<a:tailEnd type="triangle"/></a:ln>`, name: 'Canvas connector' }),
      ]),
    }),
    'wpc',
  );
  const linkedFirst = drawingRun(
    anchor({
      w: emuIn(2.0), h: emuIn(1.1), posH: { rel: 'margin', offset: 0 }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'topAndBottom' }, name: 'Linked box 1',
      graphic: shapeGraphic(wsp({ w: emuIn(2.0), h: emuIn(1.1), fill: solid(rgb('EEF3FA')), line: `<a:ln w="9525">${solid(rgb('2F5496'))}</a:ln>`, style: '', txbxId: 1, text: [p({ runs: [`Linked text boxes: this text starts here and continues in the second box. ${S[0]} ${S[1]}`], after: 0 })] })),
    }),
    'wps',
  );
  const linkedSecond = drawingRun(
    anchor({
      w: emuIn(2.0), h: emuIn(1.1), posH: { rel: 'margin', offset: emuIn(3.0) }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'none' }, name: 'Linked box 2',
      graphic: shapeGraphic(wsp({ w: emuIn(2.0), h: emuIn(1.1), fill: solid(rgb('EEF3FA')), line: `<a:ln w="9525">${solid(rgb('2F5496'))}</a:ln>`, style: '', linked: [1, 1] })),
    }),
    'wps',
  );
  const chartInline = (rid, wIn, hIn) => ({ xml: `<w:r>${inline({ w: emuIn(wIn), h: emuIn(hIn), graphic: chartGraphic(rid), name: 'Chart' })}</w:r>` });
  const chartFloat = { xml: `<w:r>${anchor({ w: emuIn(2.8), h: emuIn(2.2), posH: { rel: 'margin', align: 'right' }, posV: { rel: 'paragraph', offset: 0 }, wrap: { type: 'square' }, graphic: chartGraphic('rId31'), name: 'Pie chart' })}</w:r>` };
  const body = [
    p({ style: 'Title', runs: ['Forms, shapes and charts'] }),
    p({ style: 'Heading1', runs: ['Form fields and check boxes'] }),
    form,
    p({ runs: [''] }),
    p({ style: 'Heading1', runs: ['Symbol fonts'] }),
    ...symbols,
    p({ style: 'Heading1', pageBreakBefore: true, runs: ['Shapes and text boxes'] }),
    p({ runs: ['An inline shape sits in the line: ', inlineShape, ' and the text continues after it.'] }),
    p({ runs: [floatingBox, `A floating text box on the right with square wrapping. ${text(0, 5)}`] }),
    p({ runs: ['Preset shapes with theme styles:', ...shapesRow] }),
    p({ runs: ['A red line with an oval start and an arrowhead, and a rotated rectangle:', line, rotated] }),
    p({ runs: [group, `A group of an ellipse, a picture and a text box, wrapped square on the left. ${text(3, 4)}`] }),
    p({ runs: ['A drawing canvas with two shapes and a connector:'] }),
    p({ runs: [canvas] }),
    p({ runs: ['Linked text boxes:', linkedFirst, linkedSecond] }),
    p({ style: 'Heading1', pageBreakBefore: true, runs: ['Charts'] }),
    p({ runs: [chartInline('rId30', 6.0, 3.0)] }),
    p({ runs: [chartFloat, `A pie chart floats on the right. ${text(4, 5)}`] }),
    p({ runs: [chartInline('rId32', 6.0, 2.6)] }),
    sectPr({}),
  ];
  return docx({
    title: 'Forms, shapes and charts',
    subject: 'Test document for the Word renderer: forms, symbols, shapes, charts',
    body,
    media: [['rId21', 'badge.png', PICTURES.badge]],
    charts: [['rId30', 'chart1.xml', 'bar'], ['rId31', 'chart2.xml', 'pie'], ['rId32', 'chart3.xml', 'line']],
  });
}

// ---- Document 5: VML and legacy layout ------------------------------------------------------------------------

function vmlLegacy() {
  const pict = (inner) => ({ xml: `<w:r><w:pict>${inner}</w:pict></w:r>` });
  const t75 =
    '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>';
  const t202 = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
  const abs = (left, top, w, h, extra = '') => `position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:${w}pt;height:${h}pt;z-index:251660000;mso-position-horizontal-relative:margin;mso-position-vertical-relative:paragraph${extra}`;
  const body = [
    p({ style: 'Title', runs: ['VML and legacy layout'] }),
    p({ runs: ['Shapes in the VML format of Word 97–2003, which Word still reads and writes as a fallback.'] }),
    p({ runs: ['Inline VML picture: ', pict(`${t75}<v:shape id="Picture1" o:spid="_x0000_i1025" type="#_x0000_t75" style="width:108pt;height:67.5pt"><v:imagedata r:id="rId20" o:title=""/></v:shape>`), ' and a small one ', pict('<v:shape id="Picture2" type="#_x0000_t75" style="width:24pt;height:24pt"><v:imagedata r:id="rId26" o:title=""/></v:shape>'), '.'] }),
    p({ style: 'Heading2', runs: ['Floating VML shapes'] }),
    p({
      runs: [
        pict(`<v:rect id="Rect1" style="${abs(0, 6, 90, 50)}" fillcolor="#dae3f3" strokecolor="#2f5496" strokeweight="1.5pt"><w10:wrap type="square"/></v:rect>`),
        pict(`<v:roundrect id="Round1" style="${abs(110, 6, 90, 50)}" arcsize="25%" fillcolor="#fbe4d5" strokecolor="#c55a11" strokeweight="1pt"><w10:wrap type="square"/></v:roundrect>`),
        pict(`<v:oval id="Oval1" style="${abs(220, 6, 70, 50)}" fillcolor="#e2efda" strokecolor="#548235"><w10:wrap type="square"/></v:oval>`),
        pict(`<v:line id="Line1" style="position:absolute;z-index:251661000;mso-position-horizontal-relative:margin;mso-position-vertical-relative:paragraph" from="310pt,10pt" to="460pt,50pt" strokecolor="#c00000" strokeweight="2pt"><v:stroke endarrow="block" dashstyle="dash"/></v:line>`),
        `A rectangle, a rounded rectangle, an oval and a dashed line with an arrow, all VML. ${text(0, 3)}`,
      ],
    }),
    p({ runs: [''] }),
    p({
      runs: [
        pict(`${t202}<v:shape id="TextBox1" o:spid="_x0000_s1030" type="#_x0000_t202" style="${abs(300, 0, 160, 80, ';mso-wrap-style:square;v-text-anchor:top')}" fillcolor="#fff2cc" strokecolor="#bf8f00" strokeweight="1pt"><v:textbox inset="7.2pt,3.6pt,7.2pt,3.6pt"><w:txbxContent>${p({ runs: [{ t: 'VML text box', b: true }], after: 60 })}${p({ runs: ['Text in a legacy text box wraps inside it.'], after: 0 })}</w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape>`),
        `A VML text box on the right. ${text(2, 4)}`,
      ],
    }),
    p({ runs: [pict(`<v:shape id="Picture3" type="#_x0000_t75" style="${abs(0, 0, 100, 62.5)}"><v:imagedata r:id="rId25" o:title=""/><w10:wrap type="square"/></v:shape>`), `A floating VML picture with square wrapping on the left. ${text(5, 3)}`] }),
    p({ style: 'Heading2', runs: ['Text frames and a drop cap'] }),
    p({ framePr: 'w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"', line: 806, rule: 'exact', after: 0, runs: [{ t: 'T', sz: 41, font: 'Cambria', color: '2F5496', position: 0 }] }),
    p({ runs: [`he first letter of this paragraph is a drop cap three lines high, a legacy frame that Word puts into the paragraph. ${text(1, 3)}`] }),
    p({ framePr: 'w:w="2880" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"', border: { top: ['single', 6, '7F7F7F', 1], left: ['single', 6, '7F7F7F', 4], bottom: ['single', 6, '7F7F7F', 1], right: ['single', 6, '7F7F7F', 4] }, shd: 'F2F2F2', runs: [{ t: 'A text frame aligned right, 2 inches wide, with a border. Text flows around it.', i: true }] }),
    p({ runs: [text(3, 5)] }),
    sectPr({ headers: { default: 'rId10' } }),
  ];
  const watermark = {
    xml: `<w:r><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:164.95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="SAMPLE"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>`,
  };
  return docx({
    title: 'VML and legacy layout',
    subject: 'Test document for the Word renderer: VML shapes, frames, watermark',
    body,
    media: [['rId20', 'harbor.png', PICTURES.harbor], ['rId25', 'quadrants.png', PICTURES.quadrants], ['rId26', 'icon.png', PICTURES.icon]],
    headers: [['rId10', 'header1.xml', 'header', [p({ style: 'Header', runs: [watermark, { t: 'VML test – the header carries a watermark', color: '595959' }] })]]],
  });
}

// ---- Page effects ------------------------------------------------------------------------------------------

function pageEffects() {
  // Word 2010 text effects: a gradient fill in theme colors, a thin outline and an outer shadow.
  const effects =
    '<w14:shadow w14:blurRad="50800" w14:dist="38100" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>' +
    '<w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:schemeClr w14:val="accent2"><w14:lumMod w14:val="50000"/></w14:schemeClr></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>' +
    '<w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:schemeClr w14:val="accent2"><w14:tint w14:val="66000"/><w14:satMod w14:val="160000"/></w14:schemeClr></w14:gs><w14:gs w14:pos="100000"><w14:schemeClr w14:val="accent2"><w14:shade w14:val="50000"/></w14:schemeClr></w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>';
  const fancy = (t) => ({ xml: `<w:r><w:rPr><w:b/><w:sz w:val="72"/><w:szCs w:val="72"/>${effects}</w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>` });
  const rule = { xml: '<w:r><w:pict><v:rect id="Rule1" style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict></w:r>' };
  const cellBorder = ['outset', 6, 'auto'];
  const outset = { top: cellBorder, left: cellBorder, bottom: cellBorder, right: cellBorder };
  const spaced = (rows) =>
    table({ width: [0, 'auto'], jc: 'center', borders: outset, grid: [1200, 7000], rows: rows.map(([a, b]) => ({ cells: [{ text: a, borders: outset }, { text: b, borders: outset }] })) })
      // Cell spacing, as HTML tables pasted into Word carry it.
      .replace('<w:tblBorders>', '<w:tblCellSpacing w:w="15" w:type="dxa"/><w:tblBorders>');
  const pgBorders = '<w:pgBorders w:offsetFrom="page"><w:top w:val="stars" w:sz="12" w:space="24" w:color="auto"/><w:left w:val="stars" w:sz="12" w:space="24" w:color="auto"/><w:bottom w:val="stars" w:sz="12" w:space="24" w:color="auto"/><w:right w:val="stars" w:sz="12" w:space="24" w:color="auto"/></w:pgBorders>';
  const body = [
    p({ jc: 'center', after: 600, runs: [fancy('Page effects')] }),
    p({ runs: ['The page has a color with a tiled picture over it and an art border (stars) that Word draws from pictures of its own. The title uses Word 2010 text effects: a gradient fill, an outline and a shadow.'] }),
    p({ runs: [rule] }),
    p({ runs: ['Above: a horizontal rule (Insert > Horizontal Line), a VML rectangle of width 0 that spans the text column.'] }),
    ...Array.from({ length: 7 }, (_, i) => p({ runs: [text(i, 3)] })),
    p({ style: 'Heading2', runs: ['A table with cell spacing'] }),
    spaced(Array.from({ length: 14 }, (_, i) => [String(i + 1), S[i % S.length]])),
    p({ style: 'Heading2', keepNext: true, runs: ['A keep-with-next chain longer than a page'] }),
    ...Array.from({ length: 16 }, (_, i) => p({ keepNext: true, runs: [text(i, 2)] })),
    p({ runs: ['The chain ends here. It cannot fit on one page, so Word lets it flow from where it starts instead of moving it to a new page.'] }),
    sectPr({ pgBorders }),
  ];
  return docx({
    title: 'Page effects',
    subject: 'Test document for the Word renderer: page background, art border, text effects, horizontal rule, cell spacing, long keep chains',
    body,
    media: [['rId20', 'parchment.png', PICTURES.parchment]],
    background: '<w:background w:color="FFF4D6"><v:background id="_x0000_s1025" o:bwmode="white" fillcolor="#fff4d6"><v:fill r:id="rId20" o:title="Parchment" type="tile"/></v:background></w:background>',
  });
}

// ---- Basic document ----------------------------------------------------------------------------------------

function basicDocument() {
  const line = ['single', 4, 'D3D8DE'];
  const parts = [['Part', 'Material', 'Quantity'], ['Gear z18', 'C45', '1'], ['Flange plate', 'S235JR', '2'], ['Hex bolt M8x30', '8.8', '24']];
  const [w, h] = [emuIn(10 / 3), emuIn(25 / 12)];
  const body = [
    p({ style: 'Heading1', runs: ['Sample document'] }),
    p({ runs: ['A short document with ', { t: 'bold', b: true }, ', ', { t: 'italic', i: true }, ', ', { t: 'underlined', u: 'single' }, ' and ', { t: 'colored', color: 'D9822B' }, ' text and accented letters: café, naïve, Zürich, Ø 25 ±0.1.'] }),
    p({ style: 'Heading2', runs: ['Parts list'] }),
    table({
      width: [9000, 'dxa'],
      borders: { top: line, left: line, bottom: line, right: line, insideH: line, insideV: line },
      grid: [3000, 3000, 3000],
      rows: parts.map((row, i) => ({ cells: row.map((t) => (i === 0 ? { runs: [{ t, b: true }], shd: 'F3F4F6' } : t)) })),
    }),
    p({ runs: ['Picture'] }),
    p({ runs: [{ xml: `<w:r>${inline({ w, h, graphic: picGraphic({ rid: 'rId20', w, h }) })}</w:r>` }] }),
    p({ runs: [br('page')] }),
    p({ style: 'Heading2', runs: ['Second page'] }),
    p({ runs: ['This page starts after a manual page break.'] }),
    sectPr({ w: 11906, h: 16838, margins: { top: 1417, right: 1417, bottom: 1134, left: 1417, header: 708, footer: 708 } }),
  ];
  return docx({
    title: 'Sample document',
    subject: 'Test document for the Word renderer: headings, character formatting, a table, a picture, a page break',
    body,
    media: [['rId20', 'harbor.png', PICTURES.harbor]],
  });
}

// ---- Output ------------------------------------------------------------------------------------------------

mkdirSync(ROOT, { recursive: true });
const outputs = [
  ['text-layout.docx', textLayout],
  ['sections-tables.docx', sectionsTables],
  ['floating-pictures.docx', floatingPictures],
  ['forms-shapes-charts.docx', formsShapesCharts],
  ['vml-legacy.docx', vmlLegacy],
  ['page-effects.docx', pageEffects],
  ['document.docx', basicDocument],
];
for (const [name, build] of outputs) {
  docPrId = 1;
  const data = build();
  writeFileSync(path.join(ROOT, name), data);
  console.log(`test/fixtures/${name}  ${(data.length / 1024).toFixed(1)} KB`);
}

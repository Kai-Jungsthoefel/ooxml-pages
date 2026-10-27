// SPDX-License-Identifier: MIT
// Writes PowerPoint samples that exercise the presentation renderer: preset shapes, lines and arrowheads,
// fills, rotation and groups (test/fixtures/shapes.pptx), text formatting, tables and charts in 4:3
// (test/fixtures/text-tables-charts.pptx) and the basics: a slide background, a picture, placeholders that
// inherit from layout and master, a group (test/fixtures/presentation.pptx). Everything here is written from
// scratch, so the output can be published as CC0-1.0. Usage: node scripts/generate-pptx-samples.mjs
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
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(fn(x, y), y * (w * 4 + 1) + 1 + x * 4);
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

// A small checkered tile and a soft landscape-like picture, both drawn here.
const tilePng = png(32, 32, (x, y) => ((x >> 3) + (y >> 3)) % 2 ? [242, 154, 56, 255] : [255, 244, 232, 255]);
const picturePng = png(320, 200, (x, y) => {
  const sky = y < 120 - 25 * Math.sin(x / 40);
  const sun = (x - 230) ** 2 + (y - 55) ** 2 < 28 ** 2;
  if (sun) return [255, 214, 102, 255];
  return sky ? [120 + y / 2, 170 + y / 3, 230, 255] : [60 + x / 8, 140 - (y - 120) / 2, 90, 255];
});

// ---- XML helpers ------------------------------------------------------------------------------------------

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';

const xml = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
/** Points to EMU. */
const pt = (n) => Math.round(n * 12700);
const rels = (list) =>
  xml(`<Relationships xmlns="${PKG}">${list.map(([id, type, target, mode]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${esc(target)}"${mode ? ` TargetMode="${mode}"` : ''}/>`).join('')}</Relationships>`);

const xfrm = ([x, y, w, h], { rot, flipH, flipV } = {}) =>
  `<a:xfrm${rot ? ` rot="${Math.round(rot * 60000)}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}><a:off x="${pt(x)}" y="${pt(y)}"/><a:ext cx="${pt(w)}" cy="${pt(h)}"/></a:xfrm>`;
const scheme = (name, mods = '') => `<a:schemeClr val="${name}">${mods}</a:schemeClr>`;
const rgb = (hex, mods = '') => `<a:srgbClr val="${hex}">${mods}</a:srgbClr>`;
const solid = (color) => `<a:solidFill>${color}</a:solidFill>`;
const geom = (prst, av = {}) => `<a:prstGeom prst="${prst}"><a:avLst>${Object.entries(av).map(([k, v]) => `<a:gd name="${k}" fmla="val ${v}"/>`).join('')}</a:avLst></a:prstGeom>`;
// The usual style of shapes drawn in PowerPoint: theme line and fill with the accent color, light text.
const shapeStyle = (accent = 'accent1') =>
  `<p:style><a:lnRef idx="2">${scheme(accent, '<a:shade val="50000"/>')}</a:lnRef><a:fillRef idx="1">${scheme(accent)}</a:fillRef><a:effectRef idx="0">${scheme(accent)}</a:effectRef><a:fontRef idx="minor">${scheme('lt1')}</a:fontRef></p:style>`;
const lineStyle = `<p:style><a:lnRef idx="1">${scheme('accent1')}</a:lnRef><a:fillRef idx="0">${scheme('accent1')}</a:fillRef><a:effectRef idx="0">${scheme('accent1')}</a:effectRef><a:fontRef idx="minor">${scheme('tx1')}</a:fontRef></p:style>`;

/** Runs: strings or { t, sz, b, i, u, strike, baseline, color, highlight, font, link }. */
function run(r) {
  if (typeof r === 'string') r = { t: r };
  if (r.br) return '<a:br><a:rPr lang="en-US"/></a:br>';
  const attrs = [
    'lang="en-US"',
    r.sz && `sz="${Math.round(r.sz * 100)}"`,
    r.b && 'b="1"',
    r.i && 'i="1"',
    r.u && `u="${r.u}"`,
    r.strike && `strike="${r.strike}"`,
    r.baseline && `baseline="${r.baseline}"`,
    r.cap && `cap="${r.cap}"`,
  ].filter(Boolean);
  const inner = [
    r.color && solid(r.color),
    r.highlight && `<a:highlight>${r.highlight}</a:highlight>`,
    r.font && `<a:latin typeface="${esc(r.font)}"/>`,
    r.link && `<a:hlinkClick r:id="${r.link}"/>`,
  ].filter(Boolean);
  return `<a:r><a:rPr ${attrs.join(' ')}${inner.length ? `>${inner.join('')}</a:rPr>` : '/>'}<a:t>${esc(r.t)}</a:t></a:r>`;
}

/** Paragraphs: strings or { runs, algn, lvl, marL, indent, bullet, num, lnSpc, spcBef, spcAft, tabs }. */
function para(p) {
  if (typeof p === 'string') p = { runs: [p] };
  const attrs = [
    p.lvl && `lvl="${p.lvl}"`,
    p.marL !== undefined && `marL="${pt(p.marL)}"`,
    p.indent !== undefined && `indent="${pt(p.indent)}"`,
    p.algn && `algn="${p.algn}"`,
  ].filter(Boolean);
  const inner = [
    p.lnSpc && `<a:lnSpc><a:spcPct val="${p.lnSpc * 1000}"/></a:lnSpc>`,
    p.spcBef !== undefined && `<a:spcBef><a:spcPts val="${p.spcBef * 100}"/></a:spcBef>`,
    p.spcAft !== undefined && `<a:spcAft><a:spcPts val="${p.spcAft * 100}"/></a:spcAft>`,
    p.buColor && `<a:buClr>${p.buColor}</a:buClr>`,
    p.buSize && `<a:buSzPct val="${p.buSize * 1000}"/>`,
    p.bullet === false && '<a:buNone/>',
    typeof p.bullet === 'string' && `<a:buFont typeface="Arial"/><a:buChar char="${esc(p.bullet)}"/>`,
    p.wingding && `<a:buFont typeface="Wingdings" pitchFamily="2" charset="2"/><a:buChar char="${esc(p.wingding)}"/>`,
    p.num && `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${p.num}"/>`,
    p.tabs && `<a:tabLst>${p.tabs.map(([pos, algn]) => `<a:tab pos="${pt(pos)}" algn="${algn}"/>`).join('')}</a:tabLst>`,
  ].filter(Boolean);
  const pPr = attrs.length || inner.length ? `<a:pPr ${attrs.join(' ')}${inner.length ? `>${inner.join('')}</a:pPr>` : '/>'}` : '';
  return `<a:p>${pPr}${(p.runs ?? []).map(run).join('')}<a:endParaRPr lang="en-US"${p.endSz ? ` sz="${p.endSz * 100}"` : ''}/></a:p>`;
}

/** body: bodyPr attributes, or [attributes, child elements] for autofit settings. */
function txBody(paras, body = '') {
  const [attrs, inner] = Array.isArray(body) ? body : [body, ''];
  return `<p:txBody><a:bodyPr ${attrs}${inner ? `>${inner}</a:bodyPr>` : '/>'}<a:lstStyle/>${paras.map(para).join('')}</p:txBody>`;
}

let nextId = 2;
function sp({ name = 'Shape', box, prst = 'rect', av, fill, line, style, text, body, at = {}, txBox = false, cust }) {
  const id = nextId++;
  const spPr = `<p:spPr>${xfrm(box, at)}${cust ?? geom(prst, av)}${fill ?? ''}${line ?? ''}</p:spPr>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)} ${id}"/><p:cNvSpPr${txBox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>${spPr}${style ?? ''}${text ? txBody(text, body) : ''}</p:sp>`;
}

function pic({ name = 'Picture', box, rid }) {
  const id = nextId++;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)} ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(box)}${geom('rect')}</p:spPr></p:pic>`;
}

function cxn({ box, prst = 'straightConnector1', line, at = {}, av }) {
  const id = nextId++;
  return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm(box, at)}${geom(prst, av)}${line ?? ''}</p:spPr>${lineStyle}</p:cxnSp>`;
}

function label(box, text, size = 10, extra = {}) {
  return sp({ name: 'Text box', box, txBox: true, text: [{ runs: [{ t: text, sz: size, ...extra }], algn: extra.algn }], body: 'wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"' });
}

function title(text, layout = 'title') {
  const id = nextId++;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${layout}"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody([text])}</p:sp>`;
}

const emptyTree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const slideXml = (shapes, { hidden = false, bg = '' } = {}) =>
  xml(`<p:sld ${NS}${hidden ? ' show="0"' : ''}><p:cSld>${bg}<p:spTree>${emptyTree}${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);

// ---- Theme, master, layouts --------------------------------------------------------------------------------

const COLORS = [['dk1', '1B1E23'], ['lt1', 'FFFFFF'], ['dk2', '2D3A4A'], ['lt2', 'EEF1F4'], ['accent1', '2F6DB5'], ['accent2', 'E2783A'], ['accent3', '3E9C5C'], ['accent4', 'D9A82B'], ['accent5', '8A5BB8'], ['accent6', 'C94F64'], ['hlink', '1F64C8'], ['folHlink', '7B5AA6']];

function theme() {
  const fills = `${solid(scheme('phClr'))}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${scheme('phClr', '<a:tint val="70000"/>')}</a:gs><a:gs pos="100000">${scheme('phClr', '<a:shade val="90000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>${solid(scheme('phClr', '<a:shade val="80000"/>'))}`;
  const lines = [6350, 12700, 19050].map((w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid(scheme('phClr'))}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join('');
  const bgFills = `${solid(scheme('phClr'))}${solid(scheme('phClr', '<a:tint val="95000"/>'))}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${scheme('phClr', '<a:tint val="95000"/>')}</a:gs><a:gs pos="100000">${scheme('phClr', '<a:shade val="85000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
  const font = (face) => `<a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  return xml(
    `<a:theme xmlns:a="${A}" name="Sample"><a:themeElements><a:clrScheme name="Sample">${COLORS.map(([k, v]) => `<a:${k}>${rgb(v)}</a:${k}>`).join('')}</a:clrScheme>` +
      `<a:fontScheme name="Sample"><a:majorFont>${font('Calibri Light')}</a:majorFont><a:minorFont>${font('Calibri')}</a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Sample"><a:fillStyleLst>${fills}</a:fillStyleLst><a:lnStyleLst>${lines}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${bgFills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
  );
}

function placeholder(id, name, ph, box, text) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${box ? xfrm(box) : ''}</p:spPr>${txBody([text])}</p:sp>`;
}

function master(w, h) {
  const level = (n, size) =>
    `<a:lvl${n}pPr marL="${pt(18 + (n - 1) * 27)}" indent="${pt(-18)}" algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${n === 1 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="${size}">${solid(scheme('tx1'))}<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl${n}pPr>`;
  const txStyles =
    `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/><a:defRPr sz="3600">${solid(scheme('tx1'))}<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
    `<p:bodyStyle>${[2400, 2000, 1800, 1600, 1600].map((s, i) => level(i + 1, s)).join('')}</p:bodyStyle>` +
    `<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800">${solid(scheme('tx1'))}<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles>`;
  // A thin accent bar at the top belongs to the master and shows on every slide.
  const bar = `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Accent bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>${xfrm([0, 0, w, 6])}${geom('rect')}${solid(scheme('accent1'))}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
  const tree = `${emptyTree}${bar}${placeholder(2, 'Title', '<p:ph type="title"/>', [36, 18, w - 72, 50], 'Title')}${placeholder(3, 'Text', '<p:ph type="body" idx="1"/>', [36, 80, w - 72, h - 130], 'Text')}${placeholder(4, 'Slide number', '<p:ph type="sldNum" sz="quarter" idx="12"/>', [w - 96, h - 30, 60, 20], '‹#›')}`;
  return xml(
    `<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001">${scheme('bg1')}</p:bgRef></p:bg><p:spTree>${tree}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>${txStyles}</p:sldMaster>`,
  );
}

function layouts() {
  const titleOnly = `${emptyTree}${placeholder(2, 'Title', '<p:ph type="title"/>', null, 'Title')}`;
  return [
    xml(`<p:sldLayout ${NS} type="titleOnly" preserve="1"><p:cSld name="Title only"><p:spTree>${titleOnly}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`),
    xml(`<p:sldLayout ${NS} type="obj" preserve="1"><p:cSld name="Title and content"><p:spTree>${titleOnly}${placeholder(3, 'Content', '<p:ph idx="1"/>', null, 'Text')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`),
  ];
}

/**
 * A complete package. slides: [{ xml, rels: [[id, type, target]], layout: 1|2, parts: [[name, content, contentType]] }].
 */
function presentation({ w, h, slides, media = [], tableStyles = '' }) {
  const overrides = [
    ['/ppt/presentation.xml', `${CT}.presentationml.presentation.main+xml`],
    ['/ppt/slideMasters/slideMaster1.xml', `${CT}.presentationml.slideMaster+xml`],
    ['/ppt/slideLayouts/slideLayout1.xml', `${CT}.presentationml.slideLayout+xml`],
    ['/ppt/slideLayouts/slideLayout2.xml', `${CT}.presentationml.slideLayout+xml`],
    ['/ppt/theme/theme1.xml', `${CT}.theme+xml`],
    ['/ppt/tableStyles.xml', `${CT}.presentationml.tableStyles+xml`],
    ...slides.map((_, i) => [`/ppt/slides/slide${i + 1}.xml`, `${CT}.presentationml.slide+xml`]),
    ...slides.flatMap((s) => (s.parts ?? []).filter((p) => p[2]).map(([name, , type]) => [`/${name}`, type])),
  ];
  const contentTypes = xml(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('') +
      '</Types>',
  );
  const presRels = [['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'], ['rId2', 'theme', 'theme/theme1.xml'], ['rId3', 'tableStyles', 'tableStyles.xml'], ...slides.map((_, i) => [`rId${10 + i}`, 'slide', `slides/slide${i + 1}.xml`])];
  const pres = xml(
    `<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${pt(w)}" cy="${pt(h)}"/><p:notesSz cx="6858000" cy="9144000"/>` +
      `<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800">${solid(scheme('tx1'))}<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>`,
  );
  return zip([
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', rels([['rId1', 'officeDocument', 'ppt/presentation.xml']])],
    ['ppt/presentation.xml', pres],
    ['ppt/_rels/presentation.xml.rels', rels(presRels)],
    ['ppt/tableStyles.xml', xml(`<a:tblStyleLst xmlns:a="${A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${tableStyles}</a:tblStyleLst>`)],
    ['ppt/theme/theme1.xml', theme()],
    ['ppt/slideMasters/slideMaster1.xml', master(w, h)],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'slideLayout', '../slideLayouts/slideLayout2.xml'], ['rId3', 'theme', '../theme/theme1.xml']])],
    ...layouts().flatMap((content, i) => [
      [`ppt/slideLayouts/slideLayout${i + 1}.xml`, content],
      [`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']])],
    ]),
    ...slides.flatMap((s, i) => [
      [`ppt/slides/slide${i + 1}.xml`, s.xml],
      [`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels([['rId1', 'slideLayout', `../slideLayouts/slideLayout${s.layout ?? 1}.xml`], ...(s.rels ?? [])])],
      ...(s.parts ?? []).map(([name, content]) => [name, content]),
    ]),
    ...media,
  ]);
}

// ---- Deck 1: shapes ---------------------------------------------------------------------------------------

function presetSlide() {
  nextId = 2;
  const presets = [
    'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid',
    'pentagon', 'hexagon', 'octagon', 'chevron', 'homePlate', 'rightArrow', 'leftArrow', 'upArrow',
    'downArrow', 'leftRightArrow', 'upDownArrow', 'notchedRightArrow', 'star5', 'star8', 'plus', 'donut',
    'can', 'cube', 'frame', 'snip1Rect', 'round2SameRect', 'plaque', 'flowChartDocument', 'flowChartTerminator',
    'flowChartDecision', 'flowChartPredefinedProcess', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'heart', 'lightningBolt', 'pie', 'blockArc',
  ];
  const shapes = [title('Basic shapes')];
  presets.forEach((prst, i) => {
    const col = i % 8, row = Math.floor(i / 8);
    const box = [30 + col * 114, 88 + row * 88, 96, 56];
    shapes.push(sp({ name: prst, box, prst, style: shapeStyle(`accent${(i % 6) + 1}`), text: [{ runs: [{ t: String(i + 1), sz: 11, b: true }], algn: 'ctr' }], body: 'anchor="ctr"' }));
    shapes.push(label([box[0] - 4, box[1] + 60, box[2] + 8, 14], prst, 8, { algn: 'ctr' }));
  });
  return { xml: slideXml(shapes) };
}

function lineSlide() {
  nextId = 2;
  const shapes = [title('Lines, arrows and connectors')];
  const ends = ['triangle', 'stealth', 'arrow', 'diamond', 'oval'];
  ends.forEach((type, i) => {
    const y = 96 + i * 34;
    const w = [9525, 19050, 38100, 57150, 28575][i];
    shapes.push(cxn({ box: [40, y, 240, 0], line: `<a:ln w="${w}">${solid(scheme('accent1'))}<a:tailEnd type="${type}"/></a:ln>` }));
    shapes.push(cxn({ box: [320, y, 240, 0], line: `<a:ln w="19050">${solid(scheme('accent2'))}<a:headEnd type="${type}" w="sm" len="sm"/><a:tailEnd type="${type}" w="lg" len="lg"/></a:ln>` }));
    shapes.push(label([590, y - 7, 120, 14], `${type}, ${(w / 12700).toFixed(2)} pt`, 9));
  });
  const dashes = ['dash', 'sysDot', 'lgDashDot', 'sysDash', 'dashDot'];
  dashes.forEach((dash, i) => {
    const y = 290 + i * 26;
    shapes.push(cxn({ box: [40, y, 240, 0], line: `<a:ln w="28575" cap="${i % 2 ? 'rnd' : 'flat'}">${solid(scheme('accent3'))}<a:prstDash val="${dash}"/></a:ln>` }));
    shapes.push(label([290, y - 7, 120, 14], dash, 9));
  });
  // Lines in all four directions: flips decide where the arrowhead ends up.
  [[false, false], [true, false], [false, true], [true, true]].forEach(([flipH, flipV], i) => {
    shapes.push(cxn({ box: [430 + i * 60, 280, 50, 60], at: { flipH, flipV }, line: `<a:ln w="19050">${solid(scheme('accent5'))}<a:tailEnd type="triangle"/></a:ln>` }));
  });
  shapes.push(cxn({ prst: 'bentConnector3', box: [700, 100, 200, 90], line: `<a:ln w="25400">${solid(scheme('accent4'))}<a:tailEnd type="triangle"/></a:ln>` }));
  shapes.push(cxn({ prst: 'bentConnector3', box: [700, 220, 200, 90], at: { flipV: true }, av: { adj1: 25000 }, line: `<a:ln w="25400">${solid(scheme('accent6'))}<a:headEnd type="oval"/><a:tailEnd type="stealth"/></a:ln>` }));
  shapes.push(cxn({ prst: 'curvedConnector3', box: [700, 340, 200, 100], line: `<a:ln w="25400">${solid(scheme('accent1'))}<a:tailEnd type="arrow"/></a:ln>` }));
  shapes.push(cxn({ prst: 'bentConnector2', box: [430, 380, 200, 80], line: `<a:ln w="38100" cap="rnd">${solid(scheme('accent2'))}<a:tailEnd type="triangle" w="lg" len="med"/></a:ln>` }));
  return { xml: slideXml(shapes) };
}

function fillSlide() {
  nextId = 2;
  const shapes = [title('Fills, rotation and groups')];
  const grad = (ang, stops, extra = '') =>
    `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([pos, c]) => `<a:gs pos="${pos * 1000}">${c}</a:gs>`).join('')}</a:gsLst>${extra || `<a:lin ang="${ang * 60000}" scaled="0"/>`}</a:gradFill>`;
  const noLine = '<a:ln><a:noFill/></a:ln>';
  shapes.push(sp({ box: [40, 90, 150, 90], fill: grad(0, [[0, scheme('accent1')], [100, scheme('accent2')]]), line: noLine, text: [{ runs: ['Linear 0°'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  shapes.push(sp({ box: [210, 90, 150, 90], prst: 'roundRect', fill: grad(45, [[0, rgb('FFFFFF')], [50, scheme('accent3')], [100, scheme('dk2')]]), line: noLine, text: [{ runs: ['Linear 45°'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  shapes.push(sp({ box: [380, 90, 150, 90], prst: 'ellipse', fill: grad(0, [[0, rgb('FFF2CC')], [100, scheme('accent6')]], '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'), line: noLine, text: [{ runs: [{ t: 'Radial', color: scheme('dk1') }], algn: 'ctr' }], body: 'anchor="ctr"' }));
  shapes.push(sp({ box: [550, 90, 150, 90], fill: `<a:pattFill prst="pct50"><a:fgClr>${scheme('accent5')}</a:fgClr><a:bgClr>${rgb('FFFFFF')}</a:bgClr></a:pattFill>`, line: `<a:ln w="12700">${solid(scheme('accent5'))}</a:ln>` }));
  shapes.push(sp({ box: [720, 90, 150, 90], prst: 'hexagon', fill: `<a:blipFill><a:blip r:embed="rId2"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill>`, line: `<a:ln w="19050">${solid(scheme('accent2'))}</a:ln>` }));
  shapes.push(sp({ box: [40, 200, 150, 90], prst: 'ellipse', fill: `<a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, line: noLine }));
  shapes.push(sp({ box: [210, 200, 150, 90], fill: solid(scheme('accent1', '<a:alpha val="45000"/>')), line: `<a:ln w="38100">${solid(scheme('accent2', '<a:lumMod val="75000"/>'))}<a:prstDash val="dash"/></a:ln>`, text: [{ runs: ['45 % opaque'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  // Rotation and flips: text turns with the shape, flipping vertically puts it upside down.
  shapes.push(sp({ box: [400, 210, 130, 70], prst: 'rightArrow', at: { rot: 30 }, style: shapeStyle('accent3'), text: [{ runs: ['Rotated 30°'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  shapes.push(sp({ box: [560, 210, 130, 70], prst: 'rtTriangle', at: { flipH: true }, style: shapeStyle('accent4'), text: [{ runs: ['flipH'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  shapes.push(sp({ box: [720, 210, 130, 70], prst: 'chevron', at: { flipV: true }, style: shapeStyle('accent5'), text: [{ runs: ['flipV'], algn: 'ctr' }], body: 'anchor="ctr"' }));
  // A rotated group whose child coordinates are scaled (chExt twice the size) with a nested group inside.
  const child = (prst, [x, y, w, h], accent, text) => sp({ box: [x, y, w, h], prst, style: shapeStyle(accent), text: [{ runs: [text], algn: 'ctr' }], body: 'anchor="ctr"' });
  const inner = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId++}" name="Inner group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${pt(320)}" y="${pt(0)}"/><a:ext cx="${pt(200)}" cy="${pt(160)}"/><a:chOff x="0" y="0"/><a:chExt cx="${pt(100)}" cy="${pt(80)}"/></a:xfrm></p:grpSpPr>${child('ellipse', [0, 0, 45, 35], 'accent6', 'a')}${child('star5', [50, 40, 50, 40], 'accent4', 'b')}</p:grpSp>`;
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId++}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="${-8 * 60000}"><a:off x="${pt(60)}" y="${pt(320)}"/><a:ext cx="${pt(260)}" cy="${pt(80)}"/><a:chOff x="0" y="0"/><a:chExt cx="${pt(520)}" cy="${pt(160)}"/></a:xfrm></p:grpSpPr>${child('roundRect', [0, 20, 140, 120], 'accent1', 'Group')}${child('diamond', [160, 20, 140, 120], 'accent2', 'scaled')}${inner}</p:grpSp>`;
  shapes.push(group);
  // Freeform with lines, a cubic curve, a quadratic curve and an arc.
  const cust = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="40" y="20"/></a:lnTo><a:cubicBezTo><a:pt x="70" y="-20"/><a:pt x="100" y="60"/><a:pt x="130" y="20"/></a:cubicBezTo><a:quadBezTo><a:pt x="160" y="-10"/><a:pt x="180" y="40"/></a:quadBezTo><a:arcTo wR="20" hR="30" stAng="16200000" swAng="10800000"/><a:lnTo><a:pt x="200" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`;
  shapes.push(sp({ box: [420, 330, 240, 120], cust, style: shapeStyle('accent3'), text: [{ runs: ['Freeform'], algn: 'ctr' }], body: 'anchor="b"' }));
  shapes.push(sp({ box: [700, 330, 200, 120], prst: 'cloud', fill: solid(scheme('accent2', '<a:lumMod val="20000"/><a:lumOff val="80000"/>')), line: `<a:ln w="19050">${solid(scheme('accent2'))}</a:ln>`, text: [{ runs: [{ t: 'lumMod/lumOff', color: scheme('accent2', '<a:lumMod val="50000"/>') }], algn: 'ctr' }], body: 'anchor="ctr"' }));
  return { xml: slideXml(shapes), rels: [['rId2', 'image', '../media/tile.png'], ['rId3', 'image', '../media/picture.png']] };
}

function hiddenSlide() {
  nextId = 2;
  return { xml: slideXml([title('Hidden slide'), label([60, 120, 600, 30], 'This slide is hidden in the slide show.', 18)], { hidden: true }) };
}

writeOut('shapes.pptx', presentation({
  w: 960,
  h: 540,
  slides: [presetSlide(), lineSlide(), fillSlide(), hiddenSlide()],
  media: [['ppt/media/tile.png', tilePng], ['ppt/media/picture.png', picturePng]],
}));

// ---- Deck 2: text, tables, charts (4:3) ----------------------------------------------------------------------

function textSlide() {
  nextId = 2;
  const box = (b, paras, body = 'wrap="square"', name = 'Text') => sp({ name, box: b, txBox: true, text: paras, body, line: `<a:ln w="6350">${solid(scheme('bg1', '<a:lumMod val="75000"/>'))}</a:ln>` });
  const shapes = [title('Text formatting')];
  shapes.push(box([24, 76, 330, 190], [
    { runs: ['Bullet point'], bullet: '•', marL: 18, indent: -18 },
    { runs: ['Second level'], bullet: '–', marL: 42, indent: -18, lvl: 1 },
    { runs: ['Wingdings arrow'], wingding: 'Ø', marL: 42, indent: -18, lvl: 1 },
    { runs: ['Number one'], num: 'arabicPeriod', marL: 24, indent: -24 },
    { runs: ['Number two'], num: 'arabicPeriod', marL: 24, indent: -24 },
    { runs: ['Letter a'], num: 'alphaLcParenR', marL: 48, indent: -24, lvl: 1 },
    { runs: ['Letter b'], num: 'alphaLcParenR', marL: 48, indent: -24, lvl: 1 },
    { runs: ['Number three (Roman)'], num: 'romanUcPeriod', marL: 24, indent: -24 },
  ].map((p) => ({ ...p, runs: p.runs.map((t) => ({ t, sz: 16 })) })), 'wrap="square"', 'Lists'));
  shapes.push(box([370, 76, 330, 190], [
    { runs: [{ t: 'Bold', b: true, sz: 16 }, { t: ', ', sz: 16 }, { t: 'italic', i: true, sz: 16 }, { t: ', ', sz: 16 }, { t: 'underlined', u: 'sng', sz: 16 }, { t: ', ', sz: 16 }, { t: 'struck out', strike: 'sngStrike', sz: 16 }] },
    { runs: [{ t: 'H', sz: 16 }, { t: '2', sz: 16, baseline: '-25000' }, { t: 'O and x', sz: 16 }, { t: '2', sz: 16, baseline: '30000' }, { t: ' · ', sz: 16 }, { t: 'highlighted', sz: 16, highlight: rgb('FFFF00') }, { t: ' · ', sz: 16 }, { t: 'Small Caps', cap: 'small', sz: 16 }] },
    { runs: [{ t: 'Accent 2 color', sz: 16, color: scheme('accent2') }, { t: ' and ', sz: 16 }, { t: 'Georgia', font: 'Georgia', sz: 16 }, { t: ' and ', sz: 16 }, { t: 'Courier New', font: 'Courier New', sz: 16 }] },
    { runs: [{ t: 'Link to ', sz: 16 }, { t: 'example.org', sz: 16, link: 'rId2' }, { t: ' · café, Ø 25 mm, 21 °C, 3 €', sz: 16 }] },
    { runs: [{ t: 'Left', sz: 14 }], algn: 'l' },
    { runs: [{ t: 'Centered', sz: 14 }], algn: 'ctr' },
    { runs: [{ t: 'Right', sz: 14 }], algn: 'r' },
    { runs: [{ t: 'Tab stops:\tleft\tright', sz: 14 }], tabs: [[110, 'l'], [300, 'r']] },
  ], 'wrap="square"', 'Characters'));
  // Anchors and insets.
  ['t', 'ctr', 'b'].forEach((anchor, i) => {
    shapes.push(sp({ box: [24 + i * 112, 280, 104, 110], style: shapeStyle('accent1'), text: [{ runs: [{ t: `anchor ${anchor}`, sz: 14 }], algn: 'ctr' }], body: `anchor="${anchor}"` }));
  });
  shapes.push(sp({ box: [360, 280, 104, 110], style: shapeStyle('accent3'), text: [{ runs: [{ t: 'Vertical (vert)', sz: 14 }], algn: 'ctr' }], body: 'vert="vert" anchor="ctr"' }));
  shapes.push(sp({ box: [474, 280, 104, 110], style: shapeStyle('accent3'), text: [{ runs: [{ t: 'Vertical 270', sz: 14 }], algn: 'ctr' }], body: 'vert="vert270" anchor="ctr"' }));
  shapes.push(box([588, 280, 112, 110], [{ runs: [{ t: 'Shrunk by autofit so that the text fits into the box it was given.', sz: 20 }] }], ['wrap="square"', '<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>'], 'Autofit'));
  shapes.push(box([24, 404, 676, 110], [
    { runs: [{ t: 'Two columns: this paragraph flows from the left into the right column, the way PowerPoint sets text in several columns. ', sz: 14 }, { t: 'Line spacing 150 %.', sz: 14, b: true }], lnSpc: 150 },
    { runs: [{ t: 'Second paragraph with 6 pt space before it and 80 % line spacing, so the lines sit closer than usual.', sz: 14 }], lnSpc: 80, spcBef: 6 },
  ], 'wrap="square" numCol="2" spcCol="228600"', 'Columns'));
  return { xml: slideXml(shapes), rels: [['rId2', 'hyperlink', 'https://example.org/', 'External']] };
}

function tableSlide() {
  nextId = 2;
  const cell = (text, { span, rowSpan, hMerge, vMerge, fill, algn, b, anchor } = {}) =>
    `<a:tc${span ? ` gridSpan="${span}"` : ''}${rowSpan ? ` rowSpan="${rowSpan}"` : ''}${hMerge ? ' hMerge="1"' : ''}${vMerge ? ' vMerge="1"' : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${para({ runs: text ? [{ t: text, sz: 14, b }] : [], algn })}</a:txBody><a:tcPr${anchor ? ` anchor="${anchor}"` : ''}>${fill ?? ''}</a:tcPr></a:tc>`;
  const rows = [
    [cell('Part'), cell('Material'), cell('Qty'), cell('Price', { algn: 'r' })],
    [cell('Gear z18, Ø 72 mm', { rowSpan: 2, anchor: 'ctr' }), cell('C45'), cell('1', { algn: 'ctr' }), cell('€12.40', { algn: 'r' })],
    [cell('', { vMerge: true }), cell('42CrMo4'), cell('1', { algn: 'ctr' }), cell('€18.90', { algn: 'r' })],
    [cell('Flange plate'), cell('S235JR'), cell('2', { algn: 'ctr' }), cell('€7.15', { algn: 'r' })],
    [cell('M8 bolts and nuts', { span: 2 }), cell('', { hMerge: true }), cell('24', { algn: 'ctr' }), cell('€3.60', { algn: 'r' })],
    [cell('Total', { b: true }), cell(''), cell(''), cell('€49.20', { algn: 'r', b: true, fill: solid(scheme('accent4', '<a:lumMod val="40000"/><a:lumOff val="60000"/>')) })],
  ];
  const widths = [210, 150, 90, 120];
  const frame = (id, [x, y], tbl) =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pt(x)}" y="${pt(y)}"/><a:ext cx="${pt(570)}" cy="${pt(180)}"/></p:xfrm><a:graphic><a:graphicData uri="${A.replace('main', 'table')}">${tbl}</a:graphicData></a:graphic></p:graphicFrame>`;
  // The default table style is only referenced, not defined: tableStyles.xml stays empty like in many files.
  const styled = `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${widths.map((w) => `<a:gridCol w="${pt(w)}"/>`).join('')}</a:tblGrid>${rows.map((r) => `<a:tr h="${pt(26)}">${r.join('')}</a:tr>`).join('')}</a:tbl>`;
  // A second table without style: explicit borders and fills per cell.
  const border = (side, color, w = 12700) => `<a:${side} w="${w}">${solid(color)}</a:${side}>`;
  const plainCell = (text, i) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${para({ runs: [{ t: text, sz: 12 }] })}</a:txBody><a:tcPr marL="${pt(4)}" marR="${pt(4)}">${border('lnL', rgb('7F8994'))}${border('lnR', rgb('7F8994'))}${border('lnT', rgb('7F8994'))}${border('lnB', scheme('accent2'), 28575)}${i % 2 ? solid(rgb('F3F4F6')) : '<a:noFill/>'}</a:tcPr></a:tc>`;
  const plain = `<a:tbl><a:tblPr/><a:tblGrid>${[140, 140, 140].map((w) => `<a:gridCol w="${pt(w)}"/>`).join('')}</a:tblGrid>${[['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']].map((r, ri) => `<a:tr h="${pt(22)}">${r.map((t, ci) => plainCell(t, ri + ci)).join('')}</a:tr>`).join('')}</a:tbl>`;
  return { xml: slideXml([title('Tables'), frame(nextId++, [40, 90], styled), frame(nextId++, [40, 330], plain)]) };
}

function chartPart(kind) {
  const cats = ['Q1', 'Q2', 'Q3', 'Q4'];
  const series = [['North', [42, 51, 47, 63]], ['South', [35, 38, 44, 41]], ['West', [18, 27, 31, 36]]];
  const strCache = (values) => `<c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (values, fmt = 'General') => `<c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const text = (sz, color = '595959', b = false) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"${b ? ' b="1"' : ''}>${solid(rgb(color))}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  const titleXml = (t, sz = 1400) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="0">${solid(rgb('404040'))}</a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  const ser = (i, [name, values], extra = '') =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>T!$A$${i + 2}</c:f>${strCache([name])}</c:strRef></c:tx>${extra}<c:cat><c:strRef><c:f>T!$B$1:$E$1</c:f>${strCache(cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>T!$B$${i + 2}:$E$${i + 2}</c:f>${numCache(values)}</c:numRef></c:val></c:ser>`;
  const axes = (catTitle, valTitle) =>
    `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${catTitle ? titleXml(catTitle, 1000) : ''}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525">${solid(rgb('BFBFBF'))}</a:ln></c:spPr>${text(1000)}<c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx>` +
    `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525">${solid(rgb('D9D9D9'))}</a:ln></c:spPr></c:majorGridlines>${valTitle ? titleXml(valTitle, 1000) : ''}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${text(1000)}<c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const legend = (pos) => `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/>${text(1000)}</c:legend>`;
  let plot;
  if (kind === 'bar') {
    plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.map((s, i) => ser(i, s, `<c:spPr>${solid(scheme(`accent${i + 1}`))}</c:spPr><c:invertIfNegative val="0"/>`)).join('')}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes('Quarter', 'Revenue')}`;
  } else if (kind === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series
      .map((s, i) => ser(i, s, `<c:spPr><a:ln w="28575" cap="rnd">${solid(scheme(`accent${i + 1}`))}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${solid(scheme(`accent${i + 1}`))}<a:ln w="9525">${solid(scheme(`accent${i + 1}`))}</a:ln></c:spPr></c:marker>`))
      .join('')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes(null, null)}`;
  } else {
    const points = [0, 1, 2, 3].map((i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${solid(scheme(`accent${i + 1}`))}<a:ln w="19050">${solid(rgb('FFFFFF'))}</a:ln></c:spPr></c:dPt>`).join('');
    const labels = `<c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${text(1100, 'FFFFFF', true)}<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>`;
    plot = `<c:pieChart><c:varyColors val="1"/>${ser(0, series[0], points + labels)}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  const titles = { bar: 'Revenue by region', line: 'Trend', pie: 'Share of North' };
  return xml(
    `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${titleXml(titles[kind])}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend(kind === 'pie' ? 'r' : 'b')}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr>${solid(rgb('FFFFFF'))}<a:ln w="9525">${solid(rgb('D9D9D9'))}</a:ln></c:spPr>${text(1000)}</c:chartSpace>`,
  );
}

function chartSlide(first) {
  nextId = 2;
  const frame = (rid, [x, y, w, h]) =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${nextId++}" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pt(x)}" y="${pt(y)}"/><a:ext cx="${pt(w)}" cy="${pt(h)}"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const chartType = `${CT}.drawingml.chart+xml`;
  return {
    xml: slideXml([title('Charts'), frame('rId2', [24, 80, 400, 250]), frame('rId3', [440, 80, 260, 250]), frame('rId4', [24, 340, 676, 180])]),
    rels: [['rId2', 'chart', `../charts/chart${first}.xml`], ['rId3', 'chart', `../charts/chart${first + 1}.xml`], ['rId4', 'chart', `../charts/chart${first + 2}.xml`]],
    parts: [[`ppt/charts/chart${first}.xml`, chartPart('bar'), chartType], [`ppt/charts/chart${first + 1}.xml`, chartPart('pie'), chartType], [`ppt/charts/chart${first + 2}.xml`, chartPart('line'), chartType]],
  };
}

writeOut('text-tables-charts.pptx', presentation({ w: 720, h: 540, slides: [textSlide(), tableSlide(), chartSlide(1)] }));

// ---- Deck 3: the basics -------------------------------------------------------------------------------------

const noLine = '<a:ln><a:noFill/></a:ln>';

function coverSlide() {
  nextId = 2;
  const bg = `<p:bg><p:bgPr>${solid(rgb('FFF4E8'))}<a:effectLst/></p:bgPr></p:bg>`;
  return {
    xml: slideXml([
      sp({ name: 'Band', box: [0, 504, 960, 36], fill: solid(scheme('accent2')), line: noLine }),
      label([72, 168, 816, 90], 'ooxml-pages', 54, { b: true }),
      label([72, 264, 816, 54], 'Sample presentation, written by scripts/generate-pptx-samples.mjs', 24, { color: rgb('525B66') }),
    ], { bg }),
  };
}

function pictureSlide() {
  nextId = 2;
  return {
    xml: slideXml([
      title('Second slide: shapes and picture'),
      sp({ box: [48, 132, 252, 126], prst: 'roundRect', fill: solid(scheme('accent2')), line: noLine, text: [{ runs: [{ t: 'Orange area', sz: 20, color: scheme('dk1') }], algn: 'ctr' }], body: 'anchor="ctr"' }),
      sp({ box: [48, 276, 252, 126], prst: 'ellipse', fill: solid(scheme('accent3')), line: noLine }),
      pic({ box: [348, 132, 528, 330], rid: 'rId2' }),
    ]),
    rels: [['rId2', 'image', '../media/picture.png']],
  };
}

function placeholderSlide() {
  nextId = 2;
  // A content placeholder without position or text style: both come from layout and master, as in most real decks.
  const contentId = nextId++;
  const content = `<p:sp><p:nvSpPr><p:cNvPr id="${contentId}" name="Content ${contentId}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${txBody(['Position, font and bullets come from the master', 'The layout only names the placeholder', 'Below on the right: two shapes in a group'])}</p:sp>`;
  // The group scales its children from child coordinates (chOff/chExt) into its own box.
  const child = (prst, box, accent) => sp({ box, prst, style: shapeStyle(accent) });
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId++}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${pt(576)}" y="${pt(300)}"/><a:ext cx="${pt(288)}" cy="${pt(144)}"/><a:chOff x="0" y="0"/><a:chExt cx="${pt(144)}" cy="${pt(72)}"/></a:xfrm></p:grpSpPr>${child('rect', [0, 0, 69, 72], 'accent2')}${child('ellipse', [75, 0, 69, 72], 'accent3')}</p:grpSp>`;
  return { xml: slideXml([title('Third slide: placeholders and a group'), content, group]), layout: 2 };
}

writeOut('presentation.pptx', presentation({ w: 960, h: 540, slides: [coverSlide(), pictureSlide(), placeholderSlide()], media: [['ppt/media/picture.png', picturePng]] }));

function writeOut(name, data) {
  mkdirSync(ROOT, { recursive: true });
  writeFileSync(path.join(ROOT, name), data);
  console.log(`test/fixtures/${name}  ${data.length} bytes`);
}

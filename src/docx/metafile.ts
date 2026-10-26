// SPDX-License-Identifier: MIT
// Windows metafiles (EMF, WMF, and their gzip forms EMZ/WMZ), which browsers cannot show: the GDI drawing
// records are replayed on a canvas and the result becomes a PNG. Word documents use them for logos, pictures
// of embedded objects (Visio, Excel, CAD) and ActiveX controls. EMF+ records are skipped; Office writes the
// plain GDI drawing next to them.

const MAX_RECORDS = 500_000;
const MAX_SIDE = 2400;
const MAX_BITMAP_PIXELS = 40_000_000;
const MAX_INFLATED = 64 * 1024 * 1024;

type Matrix = [number, number, number, number, number, number];

interface Pen {
  style: number;
  width: number;
  color: string;
}

interface Brush {
  style: number;
  color: string;
  hatch: number;
  pattern?: CanvasPattern | null;
}

interface Font {
  height: number;
  escapement: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  face: string;
}

interface State {
  mapMode: number;
  winOrg: [number, number];
  winExt: [number, number];
  vpOrg: [number, number];
  vpExt: [number, number];
  world: Matrix;
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  fillMode: CanvasFillRule;
  pos: [number, number];
}

const BLACK_PEN: Pen = { style: 0, width: 0, color: '#000' };
const NULL_PEN: Pen = { style: 5, width: 0, color: '#000' };
const WHITE_BRUSH: Brush = { style: 0, color: '#fff', hatch: 0 };
const NULL_BRUSH: Brush = { style: 1, color: '#fff', hatch: 0 };
const DEFAULT_FONT: Font = { height: -12, escapement: 0, weight: 400, italic: false, underline: false, strike: false, face: 'Arial' };

function stock(index: number): Pen | Brush | Font | null {
  switch (index) {
    case 0:
      return WHITE_BRUSH;
    case 1:
      return { style: 0, color: '#c0c0c0', hatch: 0 };
    case 2:
      return { style: 0, color: '#808080', hatch: 0 };
    case 3:
      return { style: 0, color: '#404040', hatch: 0 };
    case 4:
      return { style: 0, color: '#000', hatch: 0 };
    case 5:
      return NULL_BRUSH;
    case 6:
      return { style: 0, width: 0, color: '#fff' };
    case 7:
      return BLACK_PEN;
    case 8:
      return NULL_PEN;
    case 10:
    case 11:
    case 16:
      return { ...DEFAULT_FONT, face: 'Courier New' };
    case 12:
    case 13:
    case 14:
    case 17:
      return DEFAULT_FONT;
    default:
      return null;
  }
}

const ROP_PATCOPY = 0x00f00021, ROP_BLACKNESS = 0x00000042, ROP_WHITENESS = 0x00ff0062;
/** Raster operations that put a source bitmap on the page as it is (masking pairs and inversions are skipped). */
const COPY_ROPS = new Set([0x00cc0020, 0x00ee0086, 0x00c000ca, 0x00330008, 0]);

const isPen = (o: unknown): o is Pen => !!o && typeof o === 'object' && 'width' in o;
const isBrush = (o: unknown): o is Brush => !!o && typeof o === 'object' && 'hatch' in o;
const isFont = (o: unknown): o is Font => !!o && typeof o === 'object' && 'face' in o;

function colorRef(v: number): string {
  const r = v & 255, g = (v >> 8) & 255, b = (v >> 16) & 255;
  return `rgb(${r},${g},${b})`;
}

function multiply(a: Matrix, b: Matrix): Matrix {
  // Applies b first, then a (x' = a(b(x))).
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

// ---- Device independent bitmaps ---------------------------------------------------------------------------

/** A DIB (header, color table, bits) as a canvas; JPEG/PNG compressed DIBs come back as decoded images. */
function decodeDib(view: DataView, bmiOffset: number, bitsOffset: number, bitsSize: number, images: Map<number, ImageBitmap | HTMLImageElement>): CanvasImageSource | null {
  const hdrSize = view.getUint32(bmiOffset, true);
  let width: number, height: number, bitCount: number, compression = 0, clrUsed = 0;
  let paletteOffset: number, paletteEntry = 4;
  if (hdrSize === 12) {
    width = view.getUint16(bmiOffset + 4, true);
    height = view.getInt16(bmiOffset + 6, true);
    bitCount = view.getUint16(bmiOffset + 10, true);
    paletteOffset = bmiOffset + 12;
    paletteEntry = 3;
  } else {
    width = view.getInt32(bmiOffset + 4, true);
    height = view.getInt32(bmiOffset + 8, true);
    bitCount = view.getUint16(bmiOffset + 14, true);
    compression = view.getUint32(bmiOffset + 16, true);
    clrUsed = view.getUint32(bmiOffset + 32, true);
    paletteOffset = bmiOffset + hdrSize;
  }
  if (compression === 4 || compression === 5) return images.get(bitsOffset) ?? null;
  const topDown = height < 0;
  height = Math.abs(height);
  if (width <= 0 || height <= 0 || width * height > MAX_BITMAP_PIXELS) return null;
  const masks = compression === 3 ? [view.getUint32(paletteOffset, true), view.getUint32(paletteOffset + 4, true), view.getUint32(paletteOffset + 8, true)] : null;
  if (masks && hdrSize === 40) paletteOffset += 12;
  const colors = bitCount <= 8 ? clrUsed || 1 << bitCount : 0;
  const palette: number[][] = [];
  for (let i = 0; i < colors; i++) {
    const o = paletteOffset + i * paletteEntry;
    if (o + 3 > view.byteLength) break;
    palette.push([view.getUint8(o + 2), view.getUint8(o + 1), view.getUint8(o)]);
  }
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  const img = ctx.createImageData(width, height);
  const out = img.data;
  const stride = Math.floor((width * bitCount + 31) / 32) * 4;
  const end = Math.min(view.byteLength, bitsOffset + (bitsSize || stride * height));
  const put = (x: number, y: number, r: number, g: number, b: number, a = 255) => {
    const row = topDown ? y : height - 1 - y;
    const i = (row * width + x) * 4;
    out[i] = r;
    out[i + 1] = g;
    out[i + 2] = b;
    out[i + 3] = a;
  };
  if (compression === 1 || compression === 2) {
    // Run-length encoded 8 or 4 bit bitmaps.
    let x = 0, y = 0, o = bitsOffset;
    const four = compression === 2;
    const color = (index: number) => palette[index] ?? [0, 0, 0];
    while (o + 1 < end && y < height) {
      const count = view.getUint8(o), value = view.getUint8(o + 1);
      o += 2;
      if (count > 0) {
        for (let k = 0; k < count && x < width; k++, x++) {
          const c = color(four ? (k % 2 ? value & 15 : value >> 4) : value);
          put(x, y, c[0], c[1], c[2]);
        }
      } else if (value === 0) {
        x = 0;
        y++;
      } else if (value === 1) break;
      else if (value === 2) {
        x += view.getUint8(o);
        y += view.getUint8(o + 1);
        o += 2;
      } else {
        for (let k = 0; k < value && x < width && o < end; k++, x++) {
          const byte = view.getUint8(o + (four ? k >> 1 : k));
          const c = color(four ? (k % 2 ? byte & 15 : byte >> 4) : byte);
          put(x, y, c[0], c[1], c[2]);
        }
        const used = four ? Math.ceil(value / 2) : value;
        o += used + (used % 2);
      }
    }
  } else {
    const shift = (mask: number) => {
      let s = 0;
      while (mask && !(mask & 1)) (mask >>>= 1), s++;
      return [s, mask];
    };
    const m = masks ? masks.map(shift) : null;
    for (let y = 0; y < height; y++) {
      const rowStart = bitsOffset + y * stride;
      if (rowStart >= end) break;
      for (let x = 0; x < width; x++) {
        switch (bitCount) {
          case 1:
          case 4:
          case 8: {
            const bitPos = x * bitCount;
            const o = rowStart + (bitPos >> 3);
            if (o >= end) break;
            const byte = view.getUint8(o);
            const index = bitCount === 8 ? byte : bitCount === 4 ? (bitPos % 8 ? byte & 15 : byte >> 4) : (byte >> (7 - (bitPos % 8))) & 1;
            const c = palette[index] ?? [0, 0, 0];
            put(x, y, c[0], c[1], c[2]);
            break;
          }
          case 16: {
            const o = rowStart + x * 2;
            if (o + 1 >= end) break;
            const v = view.getUint16(o, true);
            if (m) {
              const ch = (i: number) => (((v >>> m[i][0]) & m[i][1]) * 255) / (m[i][1] || 1);
              put(x, y, ch(0), ch(1), ch(2));
            } else put(x, y, ((v >> 10) & 31) * 8.226, ((v >> 5) & 31) * 8.226, (v & 31) * 8.226);
            break;
          }
          case 24: {
            const o = rowStart + x * 3;
            if (o + 2 >= end) break;
            put(x, y, view.getUint8(o + 2), view.getUint8(o + 1), view.getUint8(o));
            break;
          }
          case 32: {
            const o = rowStart + x * 4;
            if (o + 3 >= end) break;
            if (m) {
              const v = view.getUint32(o, true);
              const ch = (i: number) => (((v >>> m[i][0]) & m[i][1]) * 255) / (m[i][1] || 1);
              put(x, y, ch(0), ch(1), ch(2));
            } else put(x, y, view.getUint8(o + 2), view.getUint8(o + 1), view.getUint8(o), 255);
            break;
          }
        }
      }
    }
  }
  ctx.putImageData(img, 0, 0);
  return canvas;
}

// ---- Playback ---------------------------------------------------------------------------------------------

class Player {
  state: State;
  private readonly stack: State[] = [];
  private readonly objects: (Pen | Brush | Font | null)[] = [];
  private inPath = false;
  /** Canvas pixels per device unit and the device origin of the picture. */
  constructor(
    readonly ctx: CanvasRenderingContext2D,
    private readonly scale: number,
    private readonly origin: [number, number],
    private readonly pxPerMm: [number, number],
  ) {
    this.state = {
      mapMode: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1], world: [1, 0, 0, 1, 0, 0], pen: BLACK_PEN, brush: WHITE_BRUSH,
      font: DEFAULT_FONT, textColor: '#000', bkColor: '#fff', bkMode: 2, textAlign: 0, fillMode: 'evenodd', pos: [0, 0],
    };
  }

  /** Logical units to device units for the current mapping mode. */
  private pageToDevice(): Matrix {
    const s = this.state;
    let sx = 1, sy = 1;
    switch (s.mapMode) {
      case 2: // LOMETRIC, 0.1 mm
        sx = this.pxPerMm[0] / 10;
        sy = -this.pxPerMm[1] / 10;
        break;
      case 3: // HIMETRIC, 0.01 mm
        sx = this.pxPerMm[0] / 100;
        sy = -this.pxPerMm[1] / 100;
        break;
      case 4: // LOENGLISH, 0.01 inch
        sx = (this.pxPerMm[0] * 25.4) / 100;
        sy = (-this.pxPerMm[1] * 25.4) / 100;
        break;
      case 5: // HIENGLISH, 0.001 inch
        sx = (this.pxPerMm[0] * 25.4) / 1000;
        sy = (-this.pxPerMm[1] * 25.4) / 1000;
        break;
      case 6: // TWIPS
        sx = (this.pxPerMm[0] * 25.4) / 1440;
        sy = (-this.pxPerMm[1] * 25.4) / 1440;
        break;
      case 7:
      case 8:
        sx = s.winExt[0] ? s.vpExt[0] / s.winExt[0] : 1;
        sy = s.winExt[1] ? s.vpExt[1] / s.winExt[1] : 1;
        if (s.mapMode === 7) {
          const m = Math.min(Math.abs(sx), Math.abs(sy));
          sx = Math.sign(sx || 1) * m;
          sy = Math.sign(sy || 1) * m;
        }
        break;
    }
    return [sx, 0, 0, sy, s.vpOrg[0] - s.winOrg[0] * sx, s.vpOrg[1] - s.winOrg[1] * sy];
  }

  /** The full transform from logical units to canvas pixels. */
  matrix(): Matrix {
    const toCanvas: Matrix = [this.scale, 0, 0, this.scale, -this.origin[0] * this.scale, -this.origin[1] * this.scale];
    return multiply(toCanvas, multiply(this.pageToDevice(), this.state.world));
  }

  private apply(): Matrix {
    const m = this.matrix();
    this.ctx.setTransform(m[0], m[1], m[2], m[3], m[4], m[5]);
    return m;
  }

  private strokeStyle(m: Matrix): boolean {
    const pen = this.state.pen;
    if (pen.style === 5) return false;
    const unit = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
    const ctx = this.ctx;
    ctx.strokeStyle = pen.color;
    // Zero width pens are one device pixel wide whatever the scale.
    ctx.lineWidth = pen.width > 0 ? pen.width : this.scale / unit;
    const dashes: Record<number, number[]> = { 1: [6, 3], 2: [1, 2], 3: [6, 2, 1, 2], 4: [6, 2, 1, 2, 1, 2] };
    const d = dashes[pen.style & 15];
    ctx.setLineDash(d ? d.map((v) => v * ctx.lineWidth) : []);
    const cap = pen.style & 0xf00, join = pen.style & 0xf000;
    ctx.lineCap = cap === 0x100 ? 'square' : cap === 0x200 ? 'butt' : 'round';
    ctx.lineJoin = join === 0x1000 ? 'bevel' : join === 0x2000 ? 'miter' : 'round';
    return true;
  }

  private fillStyle(): boolean {
    const b = this.state.brush;
    if (b.style === 1) return false;
    if (b.style === 2) {
      this.ctx.fillStyle = this.hatch(b) ?? b.color;
      return true;
    }
    this.ctx.fillStyle = b.pattern ?? b.color;
    return true;
  }

  private hatch(b: Brush): CanvasPattern | null {
    const c = document.createElement('canvas');
    c.width = c.height = 8;
    const g = c.getContext('2d');
    if (!g) return null;
    if (this.state.bkMode === 2) {
      g.fillStyle = this.state.bkColor;
      g.fillRect(0, 0, 8, 8);
    }
    g.strokeStyle = b.color;
    g.beginPath();
    const h = b.hatch;
    if (h === 0 || h === 4) (g.moveTo(0, 4.5), g.lineTo(8, 4.5));
    if (h === 1 || h === 4) (g.moveTo(4.5, 0), g.lineTo(4.5, 8));
    if (h === 2 || h === 5) (g.moveTo(0, 0), g.lineTo(8, 8));
    if (h === 3 || h === 5) (g.moveTo(8, 0), g.lineTo(0, 8));
    g.stroke();
    return this.ctx.createPattern(c, 'repeat');
  }

  /** Draws the current path (or adds to the path being recorded). */
  private paint(fill: boolean, stroke: boolean): void {
    if (this.inPath) return;
    const m = this.apply();
    if (fill && this.fillStyle()) this.ctx.fill(this.state.fillMode);
    if (stroke && this.strokeStyle(m)) this.ctx.stroke();
  }

  /** Starts a shape; the transform must be set before its points go into the path. */
  private begin(): void {
    this.apply();
    if (!this.inPath) this.ctx.beginPath();
  }

  polygon(points: [number, number][], close: boolean, fill: boolean): void {
    if (!points.length) return;
    this.begin();
    this.ctx.moveTo(points[0][0], points[0][1]);
    for (const p of points.slice(1)) this.ctx.lineTo(p[0], p[1]);
    if (close) this.ctx.closePath();
    this.paint(fill, true);
  }

  polyPolygon(polys: [number, number][][], close: boolean, fill: boolean): void {
    this.begin();
    for (const points of polys) {
      if (!points.length) continue;
      this.ctx.moveTo(points[0][0], points[0][1]);
      for (const p of points.slice(1)) this.ctx.lineTo(p[0], p[1]);
      if (close) this.ctx.closePath();
    }
    this.paint(fill, true);
  }

  bezier(points: [number, number][], continueFromCurrent: boolean): void {
    if (!points.length) return;
    this.begin();
    let i = 0;
    if (!continueFromCurrent) {
      this.ctx.moveTo(points[0][0], points[0][1]);
      i = 1;
    } else if (!this.inPath) this.ctx.moveTo(this.state.pos[0], this.state.pos[1]);
    for (; i + 2 < points.length + 0 && i + 2 <= points.length - 1; i += 3) {
      this.ctx.bezierCurveTo(points[i][0], points[i][1], points[i + 1][0], points[i + 1][1], points[i + 2][0], points[i + 2][1]);
    }
    const last = points[points.length - 1];
    this.state.pos = [last[0], last[1]];
    this.paint(false, true);
  }

  lineTo(x: number, y: number): void {
    this.apply();
    if (this.inPath) this.ctx.lineTo(x, y);
    else {
      this.ctx.beginPath();
      this.ctx.moveTo(this.state.pos[0], this.state.pos[1]);
      this.ctx.lineTo(x, y);
      this.paint(false, true);
    }
    this.state.pos = [x, y];
  }

  moveTo(x: number, y: number): void {
    if (this.inPath) {
      this.apply();
      this.ctx.moveTo(x, y);
    }
    this.state.pos = [x, y];
  }

  rect(l: number, t: number, r: number, b: number, round = 0, roundH = 0): void {
    this.begin();
    const x = Math.min(l, r), y = Math.min(t, b), w = Math.abs(r - l), h = Math.abs(b - t);
    if (round > 0 && 'roundRect' in this.ctx) this.ctx.roundRect(x, y, w, h, [{ x: round / 2, y: (roundH || round) / 2 }]);
    else this.ctx.rect(x, y, w, h);
    this.paint(true, true);
  }

  ellipse(l: number, t: number, r: number, b: number): void {
    this.begin();
    this.ctx.ellipse((l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2, Math.abs(b - t) / 2, 0, 0, Math.PI * 2);
    this.paint(true, true);
  }

  arc(kind: 'arc' | 'chord' | 'pie', l: number, t: number, r: number, b: number, x1: number, y1: number, x2: number, y2: number): void {
    const cx = (l + r) / 2, cy = (t + b) / 2, rx = Math.abs(r - l) / 2, ry = Math.abs(b - t) / 2;
    if (!rx || !ry) return;
    const a1 = Math.atan2((y1 - cy) / ry, (x1 - cx) / rx), a2 = Math.atan2((y2 - cy) / ry, (x2 - cx) / rx);
    this.begin();
    if (kind === 'pie') this.ctx.moveTo(cx, cy);
    // GDI draws counterclockwise by default (in device space with y down).
    this.ctx.ellipse(cx, cy, rx, ry, 0, a1, a2, true);
    if (kind !== 'arc') this.ctx.closePath();
    this.paint(kind !== 'arc', true);
  }

  text(x: number, y: number, text: string, dx: number[] | null, options: number): void {
    const s = this.state;
    if (!text) return;
    if (s.textAlign & 1) [x, y] = s.pos;
    const m = this.matrix();
    const ctx = this.ctx;
    // Text is drawn upright in canvas space: font size and spacing follow the logical scale.
    const sx = Math.hypot(m[0], m[1]) || 1, sy = Math.hypot(m[2], m[3]) || 1;
    const size = Math.max(1, Math.abs(s.font.height) * sy);
    const px = m[0] * x + m[2] * y + m[4], py = m[1] * x + m[3] * y + m[5];
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.save();
    ctx.translate(px, py);
    const flipX = m[0] < 0 ? -1 : 1;
    if (s.font.escapement) ctx.rotate((-s.font.escapement / 10) * (Math.PI / 180));
    const face = s.font.face.replace(/["'\\;{}<>]/g, '') || 'Arial';
    ctx.font = `${s.font.italic ? 'italic ' : ''}${s.font.weight >= 600 ? '700 ' : ''}${size.toFixed(2)}px "${face}", Arial, sans-serif`;
    // Positive heights are cell heights (with internal leading); text at the cell top sits ~10 % lower.
    const valign = s.textAlign & 24;
    ctx.textBaseline = valign === 24 ? 'alphabetic' : valign === 8 ? 'bottom' : 'top';
    const halign = s.textAlign & 6;
    const chars = [...text];
    const widths = dx && dx.length >= chars.length ? dx.slice(0, chars.length).map((d) => d * sx * flipX) : null;
    const total = widths ? widths.reduce((a, b) => a + b, 0) : ctx.measureText(text).width;
    let start = halign === 6 ? -total / 2 : halign === 2 ? -total : 0;
    if (s.bkMode === 2 && options & 2) {
      ctx.fillStyle = s.bkColor;
      ctx.fillRect(start, valign === 0 ? 0 : -size, total, size * 1.2);
    }
    ctx.fillStyle = s.textColor;
    if (widths) {
      for (let i = 0; i < chars.length; i++) {
        ctx.fillText(chars[i], start, 0);
        start += widths[i];
      }
    } else ctx.fillText(text, start, 0);
    if (s.font.underline || s.font.strike) {
      ctx.fillRect(halign === 6 ? -total / 2 : halign === 2 ? -total : 0, s.font.underline ? size * 0.12 : -size * 0.28, total, Math.max(1, size / 18));
    }
    ctx.restore();
    if (s.textAlign & 1) s.pos = [x + total / sx, y];
  }

  image(img: CanvasImageSource, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void {
    this.apply();
    try {
      this.ctx.imageSmoothingEnabled = true;
      // Source rectangles count rows from the top of the bitmap.
      this.ctx.drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh);
    } catch {
      // A source rectangle outside the bitmap: nothing to draw.
    }
  }

  /** A block transfer without a source bitmap: fills by its raster operation (NOP, inversions: nothing). */
  patBlt(x: number, y: number, w: number, h: number, rop: number): void {
    this.apply();
    if (rop === ROP_BLACKNESS || rop === ROP_WHITENESS) {
      this.ctx.fillStyle = rop === ROP_BLACKNESS ? '#000' : '#fff';
      this.ctx.fillRect(x, y, w, h);
    } else if (rop === ROP_PATCOPY && this.fillStyle()) this.ctx.fillRect(x, y, w, h);
  }

  // State.
  save(): void {
    this.stack.push({ ...this.state, pos: [...this.state.pos] as [number, number] });
    this.ctx.save();
  }

  restore(n: number): void {
    // Negative: that many levels back; positive: back to the n-th saved state.
    const count = n < 0 ? -n : this.stack.length - n + 1;
    for (let i = 0; i < count && this.stack.length; i++) {
      this.state = this.stack.pop()!;
      this.ctx.restore();
    }
  }

  clipRect(l: number, t: number, r: number, b: number): void {
    this.apply();
    this.ctx.beginPath();
    this.ctx.rect(Math.min(l, r), Math.min(t, b), Math.abs(r - l), Math.abs(b - t));
    this.ctx.clip();
  }

  beginPath(): void {
    this.ctx.beginPath();
    this.apply();
    this.inPath = true;
  }

  endPath(): void {
    this.inPath = false;
  }

  closeFigure(): void {
    if (this.inPath) this.ctx.closePath();
  }

  fillPath(fill: boolean, stroke: boolean): void {
    this.inPath = false;
    this.paint(fill, stroke);
  }

  clipPath(): void {
    this.inPath = false;
    this.ctx.clip(this.state.fillMode);
  }

  setObject(index: number, obj: Pen | Brush | Font | null): void {
    if (index >= 0 && index < 65536) this.objects[index] = obj;
  }

  addObject(obj: Pen | Brush | Font | null): void {
    // WMF: the lowest free slot.
    let i = 0;
    while (this.objects[i] !== undefined && this.objects[i] !== null) i++;
    if (i < 65536) this.objects[i] = obj;
  }

  select(index: number): void {
    const obj = index & 0x80000000 ? stock(index & 0x7fffffff) : this.objects[index];
    if (isPen(obj)) this.state.pen = obj;
    else if (isBrush(obj)) this.state.brush = obj;
    else if (isFont(obj)) this.state.font = obj;
  }

  remove(index: number): void {
    if (index >= 0 && index < 65536) this.objects[index] = null;
  }
}

function utf16(view: DataView, offset: number, chars: number): string {
  let out = '';
  for (let i = 0; i < chars && offset + i * 2 + 1 < view.byteLength; i++) out += String.fromCharCode(view.getUint16(offset + i * 2, true));
  return out;
}

function ansi(view: DataView, offset: number, length: number): string {
  const bytes = new Uint8Array(view.buffer, view.byteOffset + offset, Math.max(0, Math.min(length, view.byteLength - offset)));
  return new TextDecoder('windows-1252').decode(bytes);
}

function logFont(view: DataView, o: number, wide: boolean): Font {
  const face = wide ? utf16(view, o + 28, 32) : ansi(view, o + 18, 32);
  return {
    height: view.getInt32(o, true),
    escapement: view.getInt32(o + 8, true),
    weight: view.getInt32(o + 16, true),
    italic: view.getUint8(o + 20) !== 0,
    underline: view.getUint8(o + 21) !== 0,
    strike: view.getUint8(o + 22) !== 0,
    face: face.replace(/\0.*$/s, '').trim() || 'Arial',
  };
}

// ---- EMF --------------------------------------------------------------------------------------------------

async function preloadEmfImages(view: DataView): Promise<Map<number, ImageBitmap | HTMLImageElement>> {
  // JPEG and PNG compressed DIBs need decoding before the synchronous playback.
  const images = new Map<number, ImageBitmap | HTMLImageElement>();
  let o = 0;
  for (let n = 0; n < MAX_RECORDS && o + 8 <= view.byteLength; n++) {
    const type = view.getUint32(o, true), size = view.getUint32(o + 4, true);
    if (size < 8) break;
    if (type === 81 && o + 72 <= view.byteLength) {
      const bmi = o + view.getUint32(o + 48, true), bits = o + view.getUint32(o + 56, true), cb = view.getUint32(o + 60, true);
      if (bmi + 20 <= view.byteLength) {
        const compression = view.getUint32(bmi + 16, true);
        if ((compression === 4 || compression === 5) && bits + cb <= view.byteLength) {
          const blob = new Blob([new Uint8Array(view.buffer, view.byteOffset + bits, cb) as BlobPart], { type: compression === 4 ? 'image/jpeg' : 'image/png' });
          try {
            images.set(bits, await createImageBitmap(blob));
          } catch {
            // Not decodable: skipped.
          }
        }
      }
    }
    o += size;
  }
  return images;
}

/** EMF records that put something on the page (an EMF+ only file has none of them). */
const DRAWING_RECORDS = new Set([2, 3, 4, 5, 6, 7, 8, 42, 43, 44, 45, 46, 47, 54, 62, 63, 64, 76, 77, 80, 81, 83, 84, 85, 86, 87, 88, 89, 90, 91, 115]);

async function playEmf(view: DataView): Promise<HTMLCanvasElement | null> {
  if (view.byteLength < 88 || view.getUint32(0, true) !== 1 || view.getUint32(40, true) !== 0x464d4520) return null;
  const devW = view.getInt32(72, true), devH = view.getInt32(76, true), mmW = view.getInt32(80, true), mmH = view.getInt32(84, true);
  const pxPerMm: [number, number] = [devW / (mmW || 1) || 3.78, devH / (mmH || 1) || 3.78];
  // The picture is its frame (0.01 mm); the bounds of what was drawn are the fallback.
  const frame = [24, 28, 32, 36].map((k) => view.getInt32(k, true) / 100);
  let left = frame[0] * pxPerMm[0], top = frame[1] * pxPerMm[1];
  let w = (frame[2] - frame[0]) * pxPerMm[0], h = (frame[3] - frame[1]) * pxPerMm[1];
  if (!(w > 0 && h > 0)) {
    left = view.getInt32(8, true);
    top = view.getInt32(12, true);
    w = view.getInt32(16, true) - left + 1;
    h = view.getInt32(20, true) - top + 1;
  }
  if (!(w > 0 && h > 0)) return null;
  const scale = Math.min(MAX_SIDE / Math.max(w, h), 4);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  const images = await preloadEmfImages(view);
  const p = new Player(ctx, scale, [left, top], pxPerMm);
  let drawn = 0;
  const pts32 = (o: number, count: number): [number, number][] => {
    const out: [number, number][] = [];
    for (let i = 0; i < count && o + i * 8 + 7 < view.byteLength; i++) out.push([view.getInt32(o + i * 8, true), view.getInt32(o + i * 8 + 4, true)]);
    return out;
  };
  const pts16 = (o: number, count: number): [number, number][] => {
    const out: [number, number][] = [];
    for (let i = 0; i < count && o + i * 4 + 3 < view.byteLength; i++) out.push([view.getInt16(o + i * 4, true), view.getInt16(o + i * 4 + 2, true)]);
    return out;
  };
  let o = 0;
  for (let n = 0; n < MAX_RECORDS && o + 8 <= view.byteLength; n++) {
    const type = view.getUint32(o, true), size = view.getUint32(o + 4, true);
    if (size < 8 || o + size > view.byteLength) break;
    const i32 = (k: number) => view.getInt32(o + 8 + k * 4, true);
    const u32 = (k: number) => view.getUint32(o + 8 + k * 4, true);
    if (DRAWING_RECORDS.has(type)) drawn++;
    try {
      switch (type) {
        case 14: // EOF
          return drawn ? canvas : null;
        case 2: // POLYBEZIER
          p.bezier(pts32(o + 28, u32(4)), false);
          break;
        case 3: // POLYGON
        case 4: // POLYLINE
          p.polygon(pts32(o + 28, u32(4)), type === 3, type === 3);
          break;
        case 5: // POLYBEZIERTO
          p.bezier(pts32(o + 28, u32(4)), true);
          break;
        case 6: // POLYLINETO
          for (const [x, y] of pts32(o + 28, u32(4))) p.lineTo(x, y);
          break;
        case 7: // POLYPOLYLINE
        case 8: { // POLYPOLYGON
          const polys = u32(4), total = u32(5);
          const counts = Array.from({ length: Math.min(polys, 100000) }, (_, k) => view.getUint32(o + 32 + k * 4, true));
          const all = pts32(o + 32 + polys * 4, total);
          let at = 0;
          p.polyPolygon(counts.map((c) => all.slice(at, (at += c))), type === 8, type === 8);
          break;
        }
        case 85: // POLYBEZIER16
          p.bezier(pts16(o + 28, u32(4)), false);
          break;
        case 86: // POLYGON16
        case 87: // POLYLINE16
          p.polygon(pts16(o + 28, u32(4)), type === 86, type === 86);
          break;
        case 88: // POLYBEZIERTO16
          p.bezier(pts16(o + 28, u32(4)), true);
          break;
        case 89: // POLYLINETO16
          for (const [x, y] of pts16(o + 28, u32(4))) p.lineTo(x, y);
          break;
        case 90: // POLYPOLYLINE16
        case 91: { // POLYPOLYGON16
          const polys = u32(4), total = u32(5);
          const counts = Array.from({ length: Math.min(polys, 100000) }, (_, k) => view.getUint32(o + 32 + k * 4, true));
          const all = pts16(o + 32 + polys * 4, total);
          let at = 0;
          p.polyPolygon(counts.map((c) => all.slice(at, (at += c))), type === 91, type === 91);
          break;
        }
        case 9: // SETWINDOWEXTEX
          p.state.winExt = [i32(0), i32(1)];
          break;
        case 10: // SETWINDOWORGEX
          p.state.winOrg = [i32(0), i32(1)];
          break;
        case 11: // SETVIEWPORTEXTEX
          p.state.vpExt = [i32(0), i32(1)];
          break;
        case 12: // SETVIEWPORTORGEX
          p.state.vpOrg = [i32(0), i32(1)];
          break;
        case 17: // SETMAPMODE
          p.state.mapMode = u32(0);
          break;
        case 18: // SETBKMODE
          p.state.bkMode = u32(0);
          break;
        case 19: // SETPOLYFILLMODE
          p.state.fillMode = u32(0) === 2 ? 'nonzero' : 'evenodd';
          break;
        case 22: // SETTEXTALIGN
          p.state.textAlign = u32(0);
          break;
        case 24: // SETTEXTCOLOR
          p.state.textColor = colorRef(u32(0));
          break;
        case 25: // SETBKCOLOR
          p.state.bkColor = colorRef(u32(0));
          break;
        case 27: // MOVETOEX
          p.moveTo(i32(0), i32(1));
          break;
        case 30: // INTERSECTCLIPRECT
          p.clipRect(i32(0), i32(1), i32(2), i32(3));
          break;
        case 33: // SAVEDC
          p.save();
          break;
        case 34: // RESTOREDC
          p.restore(i32(0));
          break;
        case 35: // SETWORLDTRANSFORM
        case 36: { // MODIFYWORLDTRANSFORM
          const f = (k: number) => view.getFloat32(o + 8 + k * 4, true);
          const m: Matrix = [f(0), f(1), f(2), f(3), f(4), f(5)];
          if (type === 35) p.state.world = m;
          else {
            const mode = u32(6);
            if (mode === 1) p.state.world = [1, 0, 0, 1, 0, 0];
            else if (mode === 2) p.state.world = multiply(p.state.world, m);
            else if (mode === 3) p.state.world = multiply(m, p.state.world);
            else if (mode === 4) p.state.world = m;
          }
          break;
        }
        case 37: // SELECTOBJECT
          p.select(u32(0));
          break;
        case 38: { // CREATEPEN
          p.setObject(u32(0), { style: u32(1), width: i32(2), color: colorRef(u32(4)) });
          break;
        }
        case 95: { // EXTCREATEPEN
          p.setObject(u32(0), { style: u32(5), width: u32(6), color: colorRef(u32(8)) });
          break;
        }
        case 39: { // CREATEBRUSHINDIRECT
          p.setObject(u32(0), { style: u32(1), color: colorRef(u32(2)), hatch: u32(3) });
          break;
        }
        case 94: { // CREATEDIBPATTERNBRUSHPT
          const img = decodeDib(view, o + u32(2), o + u32(4), u32(5), images);
          p.setObject(u32(0), { style: 3, color: '#000', hatch: 0, pattern: img ? ctx.createPattern(img, 'repeat') : null });
          break;
        }
        case 40: // DELETEOBJECT
          p.remove(u32(0));
          break;
        case 82: // EXTCREATEFONTINDIRECTW
          p.setObject(u32(0), logFont(view, o + 12, true));
          break;
        case 42: // ELLIPSE
          p.ellipse(i32(0), i32(1), i32(2), i32(3));
          break;
        case 43: // RECTANGLE
          p.rect(i32(0), i32(1), i32(2), i32(3));
          break;
        case 44: // ROUNDRECT
          p.rect(i32(0), i32(1), i32(2), i32(3), i32(4), i32(5));
          break;
        case 45: // ARC
        case 46: // CHORD
        case 47: // PIE
          p.arc(type === 45 ? 'arc' : type === 46 ? 'chord' : 'pie', i32(0), i32(1), i32(2), i32(3), i32(4), i32(5), i32(6), i32(7));
          break;
        case 54: // LINETO
          p.lineTo(i32(0), i32(1));
          break;
        case 59: // BEGINPATH
          p.beginPath();
          break;
        case 60: // ENDPATH
          p.endPath();
          break;
        case 61: // CLOSEFIGURE
          p.closeFigure();
          break;
        case 62: // FILLPATH
          p.fillPath(true, false);
          break;
        case 63: // STROKEANDFILLPATH
          p.fillPath(true, true);
          break;
        case 64: // STROKEPATH
          p.fillPath(false, true);
          break;
        case 67: // SELECTCLIPPATH
          p.clipPath();
          break;
        case 76: { // BITBLT
          const x = i32(4), y = i32(5), cw = i32(6), ch = i32(7), rop = u32(8);
          if (u32(20)) {
            const img = COPY_ROPS.has(rop) ? decodeDib(view, o + u32(19), o + u32(21), u32(22), images) : null;
            if (img) p.image(img, i32(9), i32(10), cw, ch, x, y, cw, ch);
          } else p.patBlt(x, y, cw, ch, rop);
          break;
        }
        case 77: // STRETCHBLT
        case 115: { // ALPHABLEND
          const x = i32(4), y = i32(5), cw = i32(6), ch = i32(7), rop = u32(8);
          const img = u32(20) && (type === 115 || COPY_ROPS.has(rop)) ? decodeDib(view, o + u32(19), o + u32(21), u32(22), images) : null;
          if (img) p.image(img, i32(9), i32(10), i32(23), i32(24), x, y, cw, ch);
          else if (type === 77 && !u32(20)) p.patBlt(x, y, cw, ch, rop);
          break;
        }
        case 80: { // SETDIBITSTODEVICE
          const img = decodeDib(view, o + u32(10), o + u32(12), u32(13), images);
          if (img) p.image(img, i32(6), i32(7), i32(8), i32(9), i32(4), i32(5), i32(8), i32(9));
          break;
        }
        case 81: { // STRETCHDIBITS
          const img = COPY_ROPS.has(u32(15)) ? decodeDib(view, o + u32(10), o + u32(12), u32(13), images) : null;
          if (img) {
            // Source y counts from the bottom for bottom-up bitmaps; full bitmaps are the common case.
            const bmh = Math.abs(view.getInt32(o + u32(10) + 8, true));
            const sy = view.getInt32(o + u32(10) + 8, true) > 0 ? bmh - i32(7) - i32(9) : i32(7);
            p.image(img, i32(6), Math.max(0, sy), i32(8), i32(9), i32(4), i32(5), i32(16), i32(17));
          }
          break;
        }
        case 83: // EXTTEXTOUTA
        case 84: { // EXTTEXTOUTW
          const t = o + 36; // EMRTEXT
          const x = view.getInt32(t, true), y = view.getInt32(t + 4, true);
          const chars = view.getUint32(t + 8, true), offString = view.getUint32(t + 12, true), options = view.getUint32(t + 16, true);
          const offDx = view.getUint32(t + 36, true);
          const count = Math.min(chars, 10000);
          const text = type === 84 ? utf16(view, o + offString, count) : ansi(view, o + offString, count);
          const dx = offDx ? Array.from({ length: count }, (_, k) => view.getInt32(o + offDx + k * 4, true)) : null;
          // Glyph indices (ETO_GLYPH_INDEX) cannot be turned back into characters without the font.
          if (!(options & 0x10)) p.text(x, y, text, dx, options);
          break;
        }
      }
    } catch {
      // A damaged record: skip it and keep drawing.
    }
    o += size;
  }
  return drawn ? canvas : null;
}

// ---- WMF --------------------------------------------------------------------------------------------------

function playWmf(view: DataView): HTMLCanvasElement | null {
  let o = 0;
  let bbox: [number, number, number, number] | null = null;
  let inch = 1440;
  if (view.byteLength >= 22 && view.getUint32(0, true) === 0x9ac6cdd7) {
    bbox = [view.getInt16(6, true), view.getInt16(8, true), view.getInt16(10, true), view.getInt16(12, true)];
    inch = view.getUint16(14, true) || 1440;
    o = 22;
  }
  if (o + 18 > view.byteLength) return null;
  const headerWords = view.getUint16(o + 2, true);
  o += headerWords * 2;
  // Without a placeable header, the window extent of the file tells the size.
  const records: [number, number, number][] = [];
  let winOrg: [number, number] = [0, 0], winExt: [number, number] | null = null;
  for (let n = 0, q = o; n < MAX_RECORDS && q + 6 <= view.byteLength; n++) {
    const size = view.getUint32(q, true) * 2, fn = view.getUint16(q + 4, true);
    if (size < 6 || q + size > view.byteLength) break;
    records.push([q, size, fn]);
    if (fn === 0x020b) winOrg = [view.getInt16(q + 8, true), view.getInt16(q + 6, true)];
    if (fn === 0x020c && !winExt) winExt = [view.getInt16(q + 8, true), view.getInt16(q + 6, true)];
    if (fn === 0) break;
    q += size;
  }
  if (!bbox) {
    if (!winExt) return null;
    bbox = [winOrg[0], winOrg[1], winOrg[0] + winExt[0], winOrg[1] + winExt[1]];
  }
  const w = Math.abs(bbox[2] - bbox[0]), h = Math.abs(bbox[3] - bbox[1]);
  if (!w || !h) return null;
  // Logical units in the bounding box map to the picture; scaled to a sensible bitmap size.
  const pxW = (w / inch) * 96, pxH = (h / inch) * 96;
  const scale = Math.min(MAX_SIDE / Math.max(pxW, pxH), 4) * (pxW / w);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  const p = new Player(ctx, scale, [Math.min(bbox[0], bbox[2]), Math.min(bbox[1], bbox[3])], [1, 1]);
  p.state.mapMode = 8;
  p.state.winOrg = [0, 0];
  p.state.winExt = [1, 1];
  p.state.vpExt = [1, 1];
  const images = new Map<number, ImageBitmap | HTMLImageElement>();
  for (const [q, , fn] of records) {
    const s16 = (k: number) => view.getInt16(q + 6 + k * 2, true);
    const u16 = (k: number) => view.getUint16(q + 6 + k * 2, true);
    const color = (k: number) => colorRef(view.getUint32(q + 6 + k * 2, true));
    const points = (k: number, count: number): [number, number][] =>
      Array.from({ length: Math.min(count, 100000) }, (_, i) => [s16(k + i * 2), s16(k + i * 2 + 1)] as [number, number]);
    try {
      switch (fn) {
        case 0x020b: // SETWINDOWORG
          p.state.winOrg = [s16(1), s16(0)];
          break;
        case 0x020c: // SETWINDOWEXT
          p.state.winExt = [s16(1), s16(0)];
          p.state.vpExt = [s16(1), s16(0)];
          break;
        case 0x0103: // SETMAPMODE
          break;
        case 0x0102: // SETBKMODE
          p.state.bkMode = u16(0);
          break;
        case 0x0209: // SETTEXTCOLOR
          p.state.textColor = color(0);
          break;
        case 0x0201: // SETBKCOLOR
          p.state.bkColor = color(0);
          break;
        case 0x0106: // SETPOLYFILLMODE
          p.state.fillMode = u16(0) === 2 ? 'nonzero' : 'evenodd';
          break;
        case 0x012e: // SETTEXTALIGN
          p.state.textAlign = u16(0);
          break;
        case 0x02fa: // CREATEPENINDIRECT
          p.addObject({ style: u16(0), width: s16(1), color: color(3) });
          break;
        case 0x02fc: // CREATEBRUSHINDIRECT
          p.addObject({ style: u16(0), color: color(1), hatch: u16(3) });
          break;
        case 0x02fb: { // CREATEFONTINDIRECT
          const f = q + 6;
          p.addObject({
            height: view.getInt16(f, true), escapement: view.getInt16(f + 4, true), weight: view.getInt16(f + 8, true), italic: view.getUint8(f + 10) !== 0,
            underline: view.getUint8(f + 11) !== 0, strike: view.getUint8(f + 12) !== 0, face: ansi(view, f + 18, 32).replace(/\0.*$/s, '').trim() || 'Arial',
          });
          break;
        }
        case 0x00f7: // CREATEPALETTE
        case 0x0142: // DIBCREATEPATTERNBRUSH
        case 0x01f9: // CREATEPATTERNBRUSH
        case 0x06ff: // CREATEREGION
          p.addObject(fn === 0x0142 ? { style: 0, color: '#808080', hatch: 0 } : null);
          break;
        case 0x012d: // SELECTOBJECT
          p.select(u16(0));
          break;
        case 0x01f0: // DELETEOBJECT
          p.remove(u16(0));
          break;
        case 0x0214: // MOVETO
          p.moveTo(s16(1), s16(0));
          break;
        case 0x0213: // LINETO
          p.lineTo(s16(1), s16(0));
          break;
        case 0x0324: // POLYGON
        case 0x0325: // POLYLINE
          p.polygon(points(1, u16(0)), fn === 0x0324, fn === 0x0324);
          break;
        case 0x0538: { // POLYPOLYGON
          const polys = u16(0);
          const counts = Array.from({ length: Math.min(polys, 10000) }, (_, k) => u16(1 + k));
          let at = 1 + polys;
          p.polyPolygon(counts.map((c) => {
            const pts = points(at, c);
            at += c * 2;
            return pts;
          }), true, true);
          break;
        }
        case 0x041b: // RECTANGLE
          p.rect(s16(3), s16(2), s16(1), s16(0));
          break;
        case 0x061c: // ROUNDRECT
          p.rect(s16(5), s16(4), s16(3), s16(2), s16(1), s16(0));
          break;
        case 0x0418: // ELLIPSE
          p.ellipse(s16(3), s16(2), s16(1), s16(0));
          break;
        case 0x0817: // ARC
        case 0x0830: // CHORD
        case 0x081a: // PIE
          p.arc(fn === 0x0817 ? 'arc' : fn === 0x0830 ? 'chord' : 'pie', s16(7), s16(6), s16(5), s16(4), s16(3), s16(2), s16(1), s16(0));
          break;
        case 0x0416: // INTERSECTCLIPRECT
          p.clipRect(s16(3), s16(2), s16(1), s16(0));
          break;
        case 0x001e: // SAVEDC
          p.save();
          break;
        case 0x0127: // RESTOREDC
          p.restore(s16(0));
          break;
        case 0x0521: { // TEXTOUT
          const len = u16(0);
          const text = ansi(view, q + 8, len);
          const at = 1 + Math.ceil(len / 2);
          p.text(s16(at + 1), s16(at), text, null, 0);
          break;
        }
        case 0x0a32: { // EXTTEXTOUT
          const y = s16(0), x = s16(1), len = u16(2), options = u16(3);
          let at = q + 6 + 8 + (options & 6 ? 8 : 0);
          const text = ansi(view, at, len);
          at += len + (len % 2);
          const size = view.getUint32(q, true) * 2;
          const dx = at + len * 2 <= q + size ? Array.from({ length: len }, (_, k) => view.getInt16(at + k * 2, true)) : null;
          p.text(x, y, text, dx, options);
          break;
        }
        case 0x0f43: { // STRETCHDIB
          const img = decodeDib(view, q + 6 + 22, q + 6 + 22 + dibHeaderSize(view, q + 6 + 22), 0, images);
          if (img) p.image(img, s16(6), s16(5), s16(4), s16(3), s16(10), s16(9), s16(8), s16(7));
          break;
        }
        case 0x0b41: { // DIBSTRETCHBLT
          const size = view.getUint32(q, true) * 2;
          if (size > 26 + 6) {
            const bmi = q + 6 + 20;
            const img = decodeDib(view, bmi, bmi + dibHeaderSize(view, bmi), 0, images);
            if (img) p.image(img, s16(5), s16(4), s16(3), s16(2), s16(9), s16(8), s16(7), s16(6));
          } else p.patBlt(s16(8), s16(7), s16(6), s16(5), view.getUint32(q + 6, true));
          break;
        }
        case 0x0940: { // DIBBITBLT
          const size = view.getUint32(q, true) * 2;
          if (size > 22 + 6) {
            const bmi = q + 6 + 16;
            const img = decodeDib(view, bmi, bmi + dibHeaderSize(view, bmi), 0, images);
            if (img) p.image(img, s16(3), s16(2), s16(5), s16(4), s16(7), s16(6), s16(5), s16(4));
          }
          break;
        }
      }
    } catch {
      // A damaged record: skip it.
    }
  }
  return canvas;
}

/** Size of a DIB's header and color table, where its pixels start. */
function dibHeaderSize(view: DataView, o: number): number {
  const hdr = view.getUint32(o, true);
  if (hdr === 12) {
    const bits = view.getUint16(o + 10, true);
    return 12 + (bits <= 8 ? (1 << bits) * 3 : 0);
  }
  const bits = view.getUint16(o + 14, true);
  const compression = view.getUint32(o + 16, true);
  const used = view.getUint32(o + 32, true);
  const colors = bits <= 8 ? used || 1 << bits : used;
  return hdr + colors * 4 + (compression === 3 && hdr === 40 ? 12 : 0);
}

async function inflate(bytes: Uint8Array): Promise<Uint8Array | null> {
  try {
    const stream = new Blob([bytes as BlobPart]).stream().pipeThrough(new DecompressionStream('gzip'));
    const reader = stream.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.length;
      if (total > MAX_INFLATED) {
        await reader.cancel();
        return null;
      }
      chunks.push(value);
    }
    const out = new Uint8Array(total);
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  } catch {
    return null;
  }
}

/** PNG of a metafile, or null when it cannot be drawn. */
export async function metafileToPng(data: Uint8Array): Promise<Blob | null> {
  let bytes = data;
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) {
    const inflated = await inflate(bytes);
    if (!inflated) return null;
    bytes = inflated;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let canvas: HTMLCanvasElement | null = null;
  if (view.byteLength >= 44 && view.getUint32(0, true) === 1 && view.getUint32(40, true) === 0x464d4520) canvas = await playEmf(view);
  else canvas = playWmf(view);
  if (!canvas) return null;
  return new Promise((resolve) => canvas!.toBlob((blob) => resolve(blob), 'image/png'));
}

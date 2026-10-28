// SPDX-License-Identifier: MIT
// Drawing operators for one PDF page. Coordinates are page pixels with the origin top left, like the DOM; the
// page's base transform turns them into points with the origin bottom left. Opacity goes through graphics state
// dictionaries, gradients through shadings.
import {
  PDFDict,
  PDFName,
  PDFOperator,
  PDFOperatorNames as Ops,
  appendBezierCurve,
  beginText,
  clip,
  clipEvenOdd,
  closePath,
  concatTransformationMatrix,
  drawObject,
  endPath,
  endText,
  fill,
  fillEvenOdd,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  setCharacterSpacing,
  setDashPattern,
  setFillingRgbColor,
  setFontAndSize,
  setGraphicsState,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingRgbColor,
  setTextMatrix,
  setTextRenderingMode,
  stroke,
  type PDFArray,
  type PDFDocument,
  type PDFFont,
  type PDFHexString,
  type PDFImage,
  type PDFPage,
  type PDFRef,
  LineCapStyle,
  LineJoinStyle,
  TextRenderingMode,
} from '@cantoo/pdf-lib';
import type { Segment } from './path.js';

/** CSS pixels (1/96 in) to points (1/72 in). */
export const PT = 0.75;

export interface Rgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

export interface StrokeStyle {
  width: number;
  dash?: number[];
  dashOffset?: number;
  cap?: 'butt' | 'round' | 'square';
  join?: 'miter' | 'round' | 'bevel';
}

export interface GradientStop {
  offset: number;
  color: Rgba;
}

export type Matrix = [number, number, number, number, number, number];

const CAPS = { butt: LineCapStyle.Butt, round: LineCapStyle.Round, square: LineCapStyle.Projecting };
const JOINS = { miter: LineJoinStyle.Miter, round: LineJoinStyle.Round, bevel: LineJoinStyle.Bevel };

export class Painter {
  private readonly ops: PDFOperator[] = [];
  private readonly states = new Map<string, PDFName>();
  private readonly fonts = new Map<PDFFont, PDFName>();
  private readonly images = new Map<PDFImage, PDFName>();
  private shadings = 0;

  constructor(
    readonly doc: PDFDocument,
    readonly page: PDFPage,
    height: number,
  ) {
    this.ops.push(concatTransformationMatrix(PT, 0, 0, -PT, 0, height * PT));
  }

  save(): void {
    this.ops.push(pushGraphicsState());
  }

  restore(): void {
    this.ops.push(popGraphicsState());
  }

  transform([a, b, c, d, e, f]: Matrix): void {
    if (a === 1 && b === 0 && c === 0 && d === 1 && e === 0 && f === 0) return;
    this.ops.push(concatTransformationMatrix(a, b, c, d, e, f));
  }

  path(segments: Segment[]): boolean {
    let any = false;
    for (const s of segments) {
      if (s.k === 'M') this.ops.push(moveTo(s.x, s.y));
      else if (s.k === 'L') this.ops.push(lineTo(s.x, s.y));
      else if (s.k === 'C') this.ops.push(appendBezierCurve(s.x1, s.y1, s.x2, s.y2, s.x, s.y));
      else this.ops.push(closePath());
      any = true;
    }
    return any;
  }

  /** Opacity for fills (ca) and strokes (CA). */
  private alpha(fillAlpha: number, strokeAlpha: number): void {
    const key = `${fillAlpha.toFixed(3)}/${strokeAlpha.toFixed(3)}`;
    let name = this.states.get(key);
    if (!name) {
      name = this.page.node.newExtGState('GS', this.doc.context.obj({ Type: 'ExtGState', ca: fillAlpha, CA: strokeAlpha }));
      this.states.set(key, name);
    }
    this.ops.push(setGraphicsState(name));
  }

  fill(color: Rgba, evenOdd = false): void {
    this.alpha(color.a, 1);
    this.ops.push(setFillingRgbColor(color.r, color.g, color.b), evenOdd ? fillEvenOdd() : fill());
  }

  private strokeStyle(style: StrokeStyle): void {
    this.ops.push(setLineWidth(style.width));
    this.ops.push(setLineCap(CAPS[style.cap ?? 'butt']), setLineJoin(JOINS[style.join ?? 'miter']));
    this.ops.push(setDashPattern(style.dash ?? [], style.dashOffset ?? 0));
  }

  stroke(color: Rgba, style: StrokeStyle): void {
    this.alpha(1, color.a);
    this.strokeStyle(style);
    this.ops.push(setStrokingRgbColor(color.r, color.g, color.b), stroke());
  }

  /** Uses the current path as clip for what follows (until restore). */
  clip(evenOdd = false): void {
    this.ops.push(evenOdd ? clipEvenOdd() : clip(), endPath());
  }

  endPath(): void {
    this.ops.push(endPath());
  }

  /** Draws an image into the rectangle, the right way up. */
  image(image: PDFImage, x: number, y: number, w: number, h: number, opacity = 1): void {
    let name = this.images.get(image);
    if (!name) {
      name = this.page.node.newXObject('Image', image.ref);
      this.images.set(image, name);
    }
    this.save();
    // Always set: the opacity of whatever was drawn before would carry over.
    this.alpha(opacity, opacity);
    this.ops.push(concatTransformationMatrix(w, 0, 0, -h, x, y + h), drawObject(name));
    this.restore();
  }

  /**
   * Fills the current clip with a gradient. Linear: coords [x0, y0, x1, y1]; radial: [x0, y0, r0, x1, y1, r1],
   * in the current user space.
   */
  gradient(kind: 'linear' | 'radial', coords: number[], stops: GradientStop[], opacity = 1): void {
    if (!stops.length) return;
    const ctx = this.doc.context;
    const sorted = [...stops].sort((a, b) => a.offset - b.offset);
    // The function must cover 0 … 1: the end colors continue to the ends.
    if (sorted[0].offset > 0) sorted.unshift({ ...sorted[0], offset: 0 });
    if (sorted[sorted.length - 1].offset < 1) sorted.push({ ...sorted[sorted.length - 1], offset: 1 });
    const shading = (space: 'DeviceRGB' | 'DeviceGray', value: (c: Rgba) => number[]): PDFRef => {
      const parts: PDFDict[] = [];
      const bounds: number[] = [];
      for (let i = 0; i + 1 < sorted.length; i++) {
        parts.push(ctx.obj({ FunctionType: 2, Domain: [0, 1], C0: value(sorted[i].color), C1: value(sorted[i + 1].color), N: 1 }));
        if (i > 0) bounds.push(Math.min(1, Math.max(0, sorted[i].offset)));
      }
      const fn =
        parts.length === 1
          ? parts[0]
          : ctx.obj({ FunctionType: 3, Domain: [0, 1], Functions: parts, Bounds: bounds, Encode: parts.flatMap(() => [0, 1]) });
      return ctx.register(ctx.obj({ ShadingType: kind === 'linear' ? 2 : 3, ColorSpace: space, Coords: coords, Function: fn, Extend: [true, true] }));
    };
    const name = this.shading(shading('DeviceRGB', (c) => [c.r, c.g, c.b]));
    const alphas = sorted.map((s) => s.color.a);
    if (alphas.every((a) => Math.abs(a - alphas[0]) < 0.004)) {
      this.alpha(alphas[0] * opacity, alphas[0] * opacity);
      this.ops.push(PDFOperator.of(Ops.ShadingFill, [name]));
      return;
    }
    // Stops of different opacity (a picture fading into a color): a soft mask made of the same gradient in gray,
    // with the opacities as gray levels. The mask lies in the user space current when the state is set, so
    // the same coordinates fit; its box only has to cover the clip.
    const mask = ctx.register(
      ctx.formXObject([PDFOperator.of(Ops.ShadingFill, [PDFName.of('Sh')])], {
        BBox: [-1e5, -1e5, 1e5, 1e5],
        Group: { Type: 'Group', S: 'Transparency', CS: 'DeviceGray' },
        Resources: { Shading: { Sh: shading('DeviceGray', (c) => [c.a]) } },
      }),
    );
    const state = this.page.node.newExtGState('GS', ctx.obj({ Type: 'ExtGState', ca: opacity, CA: opacity, SMask: { Type: 'Mask', S: 'Luminosity', G: mask } }));
    // The mask stays in the graphics state until it is restored.
    this.save();
    this.ops.push(setGraphicsState(state), PDFOperator.of(Ops.ShadingFill, [name]));
    this.restore();
  }

  private shading(ref: PDFRef): PDFName {
    const resources = this.page.node.Resources()!;
    let dict = resources.lookupMaybe(PDFName.of('Shading'), PDFDict);
    if (!dict) {
      dict = this.doc.context.obj({});
      resources.set(PDFName.of('Shading'), dict);
    }
    const name = PDFName.of(`Sh${++this.shadings}`);
    dict.set(name, ref);
    return name;
  }

  fontKey(font: PDFFont): PDFName {
    let name = this.fonts.get(font);
    if (!name) {
      name = this.page.node.newFontDictionary(font.name, font.ref);
      this.fonts.set(font, name);
    }
    return name;
  }

  /**
   * One run of text in one font, its baseline starting at (x, y). `parts` are encoded glyphs with adjustments in
   * thousandths of the font size between them (TJ), negative moving right. Text with nothing to paint (its area
   * filled otherwise) stays as invisible text, to be selected and found.
   */
  text(
    font: PDFFont,
    size: number,
    x: number,
    y: number,
    parts: (PDFHexString | number)[],
    options: { color: Rgba; charSpacing?: number; fakeBold?: boolean; fakeItalic?: boolean; outline?: { color: Rgba; width: number } | null; scaleX?: number },
  ): void {
    const { color, outline } = options;
    const fillText = color.a > 0;
    const bold = !!options.fakeBold;
    const strokeColor = outline?.color ?? (bold && fillText ? color : null);
    this.alpha(fillText ? color.a : 1, strokeColor ? strokeColor.a : 1);
    if (fillText) this.ops.push(setFillingRgbColor(color.r, color.g, color.b));
    if (strokeColor) {
      this.ops.push(setStrokingRgbColor(strokeColor.r, strokeColor.g, strokeColor.b));
      this.ops.push(setLineWidth(outline?.width ?? size * 0.03), setLineJoin(LineJoinStyle.Round));
    }
    const mode = strokeColor
      ? fillText ? TextRenderingMode.FillAndOutline : TextRenderingMode.Outline
      : fillText ? TextRenderingMode.Fill : TextRenderingMode.Invisible;
    this.ops.push(
      beginText(),
      setFontAndSize(this.fontKey(font), size),
      setCharacterSpacing(options.charSpacing ?? 0),
      setTextRenderingMode(mode),
      // y flipped back, so glyphs stand upright in the page's y-down space; a slant stands in for italic.
      setTextMatrix(options.scaleX ?? 1, 0, options.fakeItalic ? 0.21 : 0, -1, x, y),
    );
    const array = this.doc.context.obj(parts) as PDFArray;
    this.ops.push(PDFOperator.of(Ops.ShowTextAdjusted, [array]), endText());
  }

  /** Writes the collected operators into the page's content stream. */
  finish(): void {
    this.page.pushOperators(...this.ops);
    this.ops.length = 0;
  }
}

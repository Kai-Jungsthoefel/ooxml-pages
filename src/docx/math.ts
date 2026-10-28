// SPDX-License-Identifier: MIT
// Office Math (OMML) as MathML, which browsers lay out themselves: fractions, scripts, radicals, n-ary operators,
// delimiters, functions, accents, bars, limits, matrices and equation arrays. Letters are italic like in Word.
import { mathFamily } from '../pptx/fonts.js';
import type { DrawingObj } from './model.js';
import { kid, kids, wa } from './xml.js';

const MATHML = 'http://www.w3.org/1998/Math/MathML';
const MAX_DEPTH = 40;
const MAX_NODES = 5000;

export const MATH_FAMILY = '"Cambria Math", "STIX Two Math", "Latin Modern Math", "Noto Sans Math", math, serif';

function m(tag: string, ...children: (Element | null)[]): Element {
  const el = document.createElementNS(MATHML, tag);
  for (const c of children) if (c) el.append(c);
  return el;
}

function token(tag: 'mi' | 'mn' | 'mo' | 'mtext', text: string, attrs: Record<string, string> = {}): Element {
  const el = document.createElementNS(MATHML, tag);
  el.textContent = text;
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
  return el;
}

/** Office stores accents as combining marks, which have no width of their own over a MathML base. */
const SPACING_ACCENTS: Record<string, string> = {
  '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜', '̄': '¯', '̅': '¯', '̆': '˘', '̇': '˙',
  '̈': '¨', '̊': '˚', '̌': 'ˇ', '⃖': '←', '⃗': '→', '⃡': '↔', '⃐': '↼', '⃑': '⇀',
};

const OPERATORS =/[+\-−=<>≤≥≠±∓×÷·∙∘∗√∑∏∐∫∬∭∮∂∇∞→←↔⇒⇐⇔∈∉∋⊂⊃⊆⊇∪∩∧∨¬∀∃∅%/\\|‖!?,;:()[\]{}⟨⟩⌈⌉⌊⌋…⋯≈≡∼≃≅∝]/;

/** Tokens of an m:r: numbers, operators, and letters one by one (italic unless plain style). */
function runTokens(r: Element, budget: { n: number }): Element[] {
  const text = kids(r)
    .filter((e) => e.localName === 't')
    .map((e) => e.textContent ?? '')
    .join('');
  const rPr = kid(r, 'rPr');
  const plain = wa(kid(rPr, 'sty'), 'val') === 'p' || !!kid(rPr, 'nor');
  const out: Element[] = [];
  const re = /(\d+(?:[.,]\d+)*)|(\s+)|(.)/gsu;
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) && budget.n-- > 0) {
    if (match[1]) out.push(token('mn', match[1]));
    else if (match[2]) out.push(token('mtext', ' '));
    else {
      const ch = match[3];
      if (OPERATORS.test(ch)) out.push(token('mo', ch === '-' ? '−' : ch));
      else out.push(token('mi', ch, plain ? { mathvariant: 'normal' } : {}));
    }
  }
  return out;
}

function prop(el: Element | null, name: string): string | null {
  return wa(kid(el, name), 'val');
}

function row(container: Element | null, depth: number, budget: { n: number }): Element {
  const out = m('mrow');
  if (!container || depth > MAX_DEPTH) return out;
  for (const c of kids(container)) {
    const converted = convert(c, depth + 1, budget);
    if (converted) out.append(converted);
  }
  return out;
}

function convert(el: Element, depth: number, budget: { n: number }): Element | null {
  if (budget.n-- <= 0 || depth > MAX_DEPTH) return null;
  const arg = (name: string) => row(kid(el, name), depth, budget);
  switch (el.localName) {
    case 'r': {
      const tokens = runTokens(el, budget);
      return tokens.length === 1 ? tokens[0] : m('mrow', ...tokens);
    }
    case 'f': {
      const type = prop(kid(el, 'fPr'), 'type');
      if (type === 'lin') return m('mrow', arg('num'), token('mo', '/'), arg('den'));
      const frac = m('mfrac', arg('num'), arg('den'));
      if (type === 'noBar') frac.setAttribute('linethickness', '0');
      return frac;
    }
    case 'sSup':
      return m('msup', arg('e'), arg('sup'));
    case 'sSub':
      return m('msub', arg('e'), arg('sub'));
    case 'sSubSup':
      return m('msubsup', arg('e'), arg('sub'), arg('sup'));
    case 'sPre':
      return m('mmultiscripts', arg('e'), m('mprescripts'), arg('sub'), arg('sup'));
    case 'rad': {
      const hide = ['1', 'on', 'true'].includes(prop(kid(el, 'radPr'), 'degHide') ?? '');
      const deg = kid(el, 'deg');
      return hide || !deg || !deg.firstElementChild ? m('msqrt', arg('e')) : m('mroot', arg('e'), arg('deg'));
    }
    case 'nary': {
      const pr = kid(el, 'naryPr');
      const chr = prop(pr, 'chr') ?? '∫';
      const op = token('mo', chr, { largeop: 'true' });
      const sub = ['1', 'on', 'true'].includes(prop(pr, 'subHide') ?? '') ? null : arg('sub');
      const sup = ['1', 'on', 'true'].includes(prop(pr, 'supHide') ?? '') ? null : arg('sup');
      const under = prop(pr, 'limLoc') === 'undOvr' || (prop(pr, 'limLoc') === null && chr !== '∫' && chr !== '∮');
      let base: Element = op;
      if (sub && sup) base = m(under ? 'munderover' : 'msubsup', op, sub, sup);
      else if (sub) base = m(under ? 'munder' : 'msub', op, sub);
      else if (sup) base = m(under ? 'mover' : 'msup', op, sup);
      return m('mrow', base, arg('e'));
    }
    case 'd': {
      const pr = kid(el, 'dPr');
      const beg = kid(pr, 'begChr') ? (prop(pr, 'begChr') ?? '') : '(';
      const end = kid(pr, 'endChr') ? (prop(pr, 'endChr') ?? '') : ')';
      const sep = kid(pr, 'sepChr') ? (prop(pr, 'sepChr') ?? '') : '|';
      const parts = kids(el).filter((c) => c.localName === 'e');
      const out = m('mrow');
      if (beg) out.append(token('mo', beg, { stretchy: 'true' }));
      parts.forEach((part, i) => {
        if (i > 0 && sep) out.append(token('mo', sep));
        out.append(row(part, depth, budget));
      });
      if (end) out.append(token('mo', end, { stretchy: 'true' }));
      return out;
    }
    case 'func':
      return m('mrow', arg('fName'), token('mo', '⁡'), arg('e'));
    case 'acc': {
      const chr = prop(kid(el, 'accPr'), 'chr') ?? '̂';
      const over = m('mover', arg('e'), token('mo', SPACING_ACCENTS[chr] ?? chr));
      over.setAttribute('accent', 'true');
      return over;
    }
    case 'bar': {
      const pos = prop(kid(el, 'barPr'), 'pos') ?? 'bot';
      return pos === 'top' ? m('mover', arg('e'), token('mo', '‾')) : m('munder', arg('e'), token('mo', '_'));
    }
    case 'groupChr': {
      const pr = kid(el, 'groupChrPr');
      const chr = prop(pr, 'chr') ?? '⏟';
      return prop(pr, 'pos') === 'top' ? m('mover', arg('e'), token('mo', chr)) : m('munder', arg('e'), token('mo', chr));
    }
    case 'limLow':
      return m('munder', arg('e'), arg('lim'));
    case 'limUpp':
      return m('mover', arg('e'), arg('lim'));
    case 'm': {
      const table = m('mtable');
      for (const mr of kids(el)) {
        if (mr.localName !== 'mr') continue;
        const tr = m('mtr');
        for (const e of kids(mr)) if (e.localName === 'e') tr.append(m('mtd', row(e, depth, budget)));
        table.append(tr);
      }
      return table;
    }
    case 'eqArr': {
      const table = m('mtable');
      for (const e of kids(el)) if (e.localName === 'e') table.append(m('mtr', m('mtd', row(e, depth, budget))));
      return table;
    }
    case 'box':
    case 'borderBox':
    case 'e':
    case 'oMath':
      return row(el.localName === 'box' || el.localName === 'borderBox' ? kid(el, 'e') : el, depth, budget);
    case 'phant':
      return m('mphantom', arg('e'));
    default:
      return null;
  }
}

/** The MathML element for an m:oMath or m:oMathPara. */
export function mathElement(omml: Element, display: boolean): Element {
  const budget = { n: MAX_NODES };
  const math = m('math');
  // The browser's own style gives <math> the generic family "math"; ours (MATH_FAMILY) is set on the container.
  (math as MathMLElement).style.fontFamily = 'inherit';
  // Display equations keep full-size fractions and limits; laid out inline so we can place them ourselves.
  if (display) math.setAttribute('displaystyle', 'true');
  const parts = omml.localName === 'oMathPara' ? kids(omml).filter((e) => e.localName === 'oMath') : [omml];
  if (parts.length > 1) {
    const table = m('mtable');
    for (const part of parts) table.append(m('mtr', m('mtd', row(part, 0, budget))));
    math.append(table);
  } else {
    math.append(row(parts[0] ?? null, 0, budget));
  }
  return math;
}

const sizes = new WeakMap<Element, Map<number, { w: number; h: number; descent: number }>>();

/** Size of an equation at a font size: width, height and the depth below the baseline. */
export function mathSize(omml: Element, _display: boolean, fontSize: number): { w: number; h: number; descent: number } {
  let bySize = sizes.get(omml);
  if (!bySize) {
    bySize = new Map();
    sizes.set(omml, bySize);
  }
  const cached = bySize.get(fontSize);
  if (cached) return cached;
  const holder = document.createElement('div');
  holder.style.cssText = `position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;line-height:normal;font-size:${fontSize}px;font-family:${mathFamily(MATH_FAMILY)}`;
  const math = mathElement(omml, _display);
  const marker = document.createElement('span');
  marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
  holder.append(math, marker);
  document.body.append(holder);
  const r = math.getBoundingClientRect();
  const base = marker.getBoundingClientRect().top;
  holder.remove();
  const size = { w: Math.max(1, r.width), h: Math.max(1, r.height), descent: Math.max(0, r.bottom - base) };
  bySize.set(fontSize, size);
  return size;
}

export function renderMath(omml: Element, obj: DrawingObj): HTMLElement {
  const wrap = document.createElement('div');
  wrap.className = 'docx-math';
  wrap.style.fontSize = `${obj.mathSize ?? 14.667}px`;
  wrap.style.fontFamily = mathFamily(MATH_FAMILY);
  wrap.append(mathElement(omml, obj.graphic.kind === 'math' && obj.graphic.display));
  return wrap;
}

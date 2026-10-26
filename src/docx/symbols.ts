// SPDX-License-Identifier: MIT
// Symbol fonts (Symbol, Wingdings, Wingdings 2, Webdings) store pictograms under Latin letters or in the private
// use area (U+F020–U+F0FF). Browsers on other systems do not have these fonts, so characters are translated to
// Unicode. Ballot boxes (☐ ☑ ☒) are drawn by the renderer itself because the fonts Word uses for them
// (MS Gothic, Segoe UI Symbol) are often missing.

const range = (from: number, chars: string): Record<number, string> => {
  const out: Record<number, string> = {};
  [...chars].forEach((ch, i) => (out[from + i] = ch));
  return out;
};

const SYMBOL: Record<number, string> = {
  ...range(0x20, ' !∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼'),
  ...range(0xa0, '€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…⏐⎯↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪'),
  ...range(0xf1, '〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭'),
};

// Wingdings after Unicode's mapping, with the common bullets and boxes replaced by characters that ordinary
// fonts have (the exact pictograms live in emoji fonts only).
const WINGDINGS: Record<number, string> = {
  ...range(0x21, '🖉✂✁👓🕭🕮🕯🕿✆🖂🖃📪📫📬📭📁📂📄🗏🗐🗄⌛🖮🖰🖲🖳🖴🖫🖬✇✍🖎✌👌👍👎☜☞☝☟🖐☺😐☹💣☠🏳🏱✈☼💧❄🕆✞🕈✠✡☪☯ॐ☸♈♉♊♋♌♍♎♏♐♑♒♓🙰🙵'),
  0x6c: '●', 0x6d: '❍', 0x6e: '■', 0x6f: '□', 0x70: '◼', 0x71: '❑', 0x72: '❒', 0x73: '⬧', 0x74: '⧫', 0x75: '◆', 0x76: '❖', 0x77: '⬥',
  0x78: '☒', 0x79: '⮹', 0x7a: '⌘', 0x7b: '🏵', 0x7c: '🏶', 0x7d: '🙶', 0x7e: '🙷',
  0x80: '⓪', ...range(0x81, '①②③④⑤⑥⑦⑧⑨⑩'), 0x8b: '⓿', ...range(0x8c, '❶❷❸❹❺❻❼❽❾❿'),
  ...range(0x96, '🙢🙠🙡🙣🙞🙜🙝🙟·•'),
  0xa0: '▪', 0xa1: '○', 0xa2: '⭕', 0xa3: '◯', 0xa4: '◉', 0xa5: '◎', 0xa6: '◯', 0xa7: '▪', 0xa8: '☐', 0xa9: '◆', 0xaa: '✦', 0xab: '★',
  0xac: '✶', 0xad: '✴', 0xae: '✹', 0xaf: '✵', 0xb0: '⯐', 0xb1: '⌖', 0xb2: '⟡', 0xb3: '⌑', 0xb4: '⯑', 0xb5: '✪', 0xb6: '✰',
  ...range(0xb7, '🕐🕑🕒🕓🕔🕕🕖🕗🕘🕙🕚🕛⮰⮱⮲⮳⮴⮵⮶⮷🙪🙫🙕🙔🙗🙖🙐🙑🙒🙓⌫⌦'),
  0xd7: '⮘', 0xd8: '➢', 0xd9: '⮙', 0xda: '⮛', 0xdb: '⮈', 0xdc: '⮊', 0xdd: '⮉', 0xde: '⮋', 0xdf: '←', 0xe0: '→', 0xe1: '↑',
  0xe2: '↓', 0xe3: '↖', 0xe4: '↗', 0xe5: '↙', 0xe6: '↘', 0xe7: '⇐', 0xe8: '➔', 0xe9: '⇑', 0xea: '⇓', 0xeb: '⇖', 0xec: '⇗',
  0xed: '⇙', 0xee: '⇘', 0xef: '⇦', 0xf0: '⇨', 0xf1: '⇧', 0xf2: '⇩', 0xf3: '⬄', 0xf4: '⇳', 0xf5: '⬀', 0xf6: '⬁', 0xf7: '⬃',
  0xf8: '⬂', 0xf9: '▭', 0xfa: '▫', 0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
};

const WINGDINGS_2: Record<number, string> = {
  ...range(0x21, '🖊🖋🖌🖍✄✀🕾🕽🗅🗆🗇🗈🗉🗊🗋🗌🗍📋🗑🗔🖵🖶🖷🖸🖭🖯🖱🖒🖓🖘🖙🖚🖛👈👉🖜🖝🖞🖟🖠🖡👆👇🖢🖣🖑'),
  0x4f: '✗', 0x50: '✓', 0x51: '☒', 0x52: '☑', 0x53: '☒', 0x54: '☒', 0x55: '☒', 0x56: '☒', 0x57: '🛇', 0x58: '⦸',
  ...range(0x69, '⓪①②③④⑤⑥⑦⑧⑨⑩⓿❶❷❸❹❺❻❼❽❾❿'),
  0x80: '☉', 0x81: '🌕', 0x82: '☽', 0x83: '☾', 0x84: '⸿', 0x85: '✝', 0x95: '⋅', 0x96: '•', 0x97: '●', 0x98: '●', 0x99: '●',
  0x9a: '◉', 0x9b: '◉', 0x9c: '◉', 0x9d: '⊙', 0x9e: '⦿', 0xa1: '◾', 0xa2: '■', 0xa3: '☐', 0xa4: '▣', 0xa5: '▣', 0xa6: '▣',
  0xa7: '▪', 0xa8: '▣', 0xad: '⬩', 0xae: '⬥', 0xaf: '◇', 0xb1: '◈', 0xb6: '⬪', 0xb7: '⬧', 0xb8: '◊', 0xba: '◖', 0xbb: '◗',
  0xbe: '⯀', 0xbf: '⯁', 0xc0: '⬟', 0xc2: '⬣', 0xc3: '⬢', 0xcf: '✦', 0xd0: '✦', 0xd2: '★', 0xd3: '★', 0xde: '✶', 0xe0: '✷',
};

const WEBDINGS: Record<number, string> = { 0x61: '✓', 0x72: '✗', 0x6e: '■', 0x3d: '=', 0x3c: '◀', 0x34: '▶', 0x35: '▲', 0x36: '▼', 0x67: '■', 0x63: '☐' };

const TABLES: Record<string, Record<number, string>> = {
  symbol: SYMBOL,
  wingdings: WINGDINGS,
  'wingdings 2': WINGDINGS_2,
  webdings: WEBDINGS,
};

export function isSymbolFont(font: string | undefined | null): boolean {
  return !!font && /^(wingdings|webdings|symbol|marlett|zapfdingbats|zapf dingbats)/i.test(font.trim());
}

/** Unicode for a code in a symbol font (plain or private use area), or null when there is no translation. */
export function symbolChar(font: string | undefined | null, code: number): string | null {
  if (!font) return null;
  const table = TABLES[font.trim().toLowerCase()];
  if (code >= 0xf000 && code <= 0xf0ff) code -= 0xf000;
  if (!table) return null;
  return table[code] ?? null;
}

/** Text of a run in a symbol font translated character by character; unknown characters stay unchanged. */
export function symbolText(font: string, text: string): { text: string; mapped: boolean } {
  let mapped = false;
  let out = '';
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    const m = symbolChar(font, code);
    if (m !== null) {
      out += m;
      mapped = true;
    } else {
      out += code >= 0xf020 && code <= 0xf0ff ? String.fromCharCode(code - 0xf000) : ch;
    }
  }
  return { text: out, mapped };
}

/** Ballot boxes the renderer draws itself. */
export const BOX_GLYPHS = new Set(['☐', '☑', '☒']);

/** Fonts in which the ballot boxes are full-width (1 em) glyphs. */
export function fullWidthBoxFont(font: string | undefined): boolean {
  return !!font && /gothic|mincho|meiryo|yu gothic|ms ui|simsun|simhei|nsimsun|malgun|gulim|batang|dotum|ming|msjh|msyh|microsoft jhenghei|microsoft yahei/i.test(font);
}

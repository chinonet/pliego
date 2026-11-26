/**
 * Lo que jsafa-pdf no ve, puesto como lo que sí ve, justo antes de leer las páginas.
 *
 * jsafa-pdf solo copia nodos de texto (su contenido tal cual, sin text-transform ni versalitas), cajas con fondo o
 * borde (todo borde sale liso) e imágenes. No ve ::before ni ::marker, ni subrayados, ni el recorte de una imagen, y
 * cada tramo de texto de un mismo nodo y renglón lo estira o lo encoge (escala horizontal) hasta el ancho medido: en un
 * párrafo justificado eso son letras estiradas. Acá se pasa todo a texto, cajas y nodos separados sin mover nada de
 * lugar (lo que movería el texto se mide y, si cambia, se deshace y se avisa). Devuelve con qué dejar todo como estaba.
 *
 * No toca el Word: solo el DOM del visor, y solo mientras se arma el PDF.
 */

export type ExportPrepOptions = {
  /** El Word tenía división automática de palabras (RenderInfo.autoHyphenation). */
  autoHyphenation: boolean;
  /** Acá se agregan los avisos para quien edita (en castellano, listos para mostrar). */
  warnings?: string[];
};

const LETTER = /\p{L}/u;
const INVISIBLE = /[\u00AD\u200B\u2060\uFEFF]/;
const INVISIBLE_SPLIT = /([\u00AD\u200B\u2060\uFEFF])/;
const SHY = "\u00AD";
const ZWSP = "\u200B";
const BLOCK = "p, li, td, th, h1, h2, h3, h4, h5, h6";
/** Si hay más letras espaciadas que esto, no se separan (el PDF tardaría demasiado). */
const MAX_SPACED_LETTERS = 150000;

// ----- Funciones puras (se prueban con node) -----

/** «docx-num-1-0 0 docx-num-1-1 3» → nombre y valor (null si el valor no está escrito). */
export function parseCounters(value: string): Array<{ name: string; value: number | null }> {
  const text = value.trim();
  if (!text || text === "none") return [];
  const tokens = text.split(/\s+/);
  const out: Array<{ name: string; value: number | null }> = [];
  for (let i = 0; i < tokens.length; i += 1) {
    const name = tokens[i]!;
    if (/^-?\d+$/.test(name)) continue;
    const next = tokens[i + 1];
    if (next !== undefined && /^-?\d+$/.test(next)) {
      out.push({ name, value: Number(next) });
      i += 1;
    } else out.push({ name, value: null });
  }
  return out;
}

export type ContentPart =
  | { kind: "text"; text: string }
  | { kind: "counter"; name: string; style: string };

function readCssString(source: string, start: number): [string, number] {
  const quote = source[start];
  let out = "";
  let i = start + 1;
  while (i < source.length) {
    const c = source[i]!;
    if (c === quote) return [out, i + 1];
    if (c === "\\") {
      const hex = /^[0-9a-fA-F]{1,6}/.exec(source.slice(i + 1, i + 7));
      if (hex) {
        const code = Number.parseInt(hex[0], 16);
        out += code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff) ? "\uFFFD" : String.fromCodePoint(code);
        i += 1 + hex[0].length;
        if (source[i] === "\r" && source[i + 1] === "\n") i += 2;
        else if (/\s/.test(source[i] ?? "")) i += 1;
        continue;
      }
      const next = source[i + 1];
      if (next === undefined) return [out, i + 1];
      if (next !== "\n") out += next;
      i += 2;
      continue;
    }
    out += c;
    i += 1;
  }
  return [out, i];
}

/** Separa por comas lo que no está entre comillas. */
function splitArgs(source: string): string[] {
  const args: string[] = [];
  let current = "";
  for (let i = 0; i < source.length; i += 1) {
    const c = source[i]!;
    if (c === '"' || c === "'") {
      const [, end] = readCssString(source, i);
      current += source.slice(i, end);
      i = end - 1;
    } else if (c === ",") {
      args.push(current.trim());
      current = "";
    } else current += c;
  }
  args.push(current.trim());
  return args;
}

/**
 * El «content» calculado de un ::before (textos entre comillas y counter()/counters()). null si no hay nada o si usa
 * algo que acá no se resuelve (attr(), comillas automáticas, imágenes).
 */
export function parseContent(value: string): ContentPart[] | null {
  const source = value.trim();
  if (!source || source === "none" || source === "normal") return null;
  const parts: ContentPart[] = [];
  let i = 0;
  while (i < source.length) {
    const c = source[i]!;
    if (/\s/.test(c)) {
      i += 1;
      continue;
    }
    // «content: "x" / "texto alternativo»: lo de después de la barra no se dibuja.
    if (c === "/") break;
    if (c === '"' || c === "'") {
      const [text, end] = readCssString(source, i);
      parts.push({ kind: "text", text });
      i = end;
      continue;
    }
    const fn = /^(counters?)\(/.exec(source.slice(i));
    if (!fn) return null;
    let depth = 0;
    let j = i + fn[0].length - 1;
    for (; j < source.length; j += 1) {
      const d = source[j]!;
      if (d === '"' || d === "'") {
        j = readCssString(source, j)[1] - 1;
        continue;
      }
      if (d === "(") depth += 1;
      if (d === ")") {
        depth -= 1;
        if (depth === 0) break;
      }
    }
    if (j >= source.length) return null;
    const args = splitArgs(source.slice(i + fn[0].length, j));
    const name = args[0] ?? "";
    if (!name) return null;
    // counters(nombre, separador, estilo): el visor usa un solo nivel por contador, así que vale como counter().
    const style = (fn[1] === "counters" ? args[2] : args[1]) ?? "decimal";
    parts.push({ kind: "counter", name, style: style || "decimal" });
    i = j + 1;
  }
  return parts;
}

const BULLETS: Record<string, string> = {
  disc: "\u2022",
  circle: "\u25E6",
  square: "\u25AA",
  "disclosure-open": "\u25BE",
  "disclosure-closed": "\u25B8",
};

/** Estilos de contador de CSS que el navegador dibuja y acá no se calculan. */
const UNSUPPORTED_STYLES = new Set([
  "armenian", "lower-armenian", "upper-armenian", "georgian", "hebrew", "cjk-decimal", "cjk-earthly-branch",
  "cjk-heavenly-stem", "cjk-ideographic", "hiragana", "hiragana-iroha", "katakana", "katakana-iroha",
  "japanese-informal", "japanese-formal", "korean-hangul-formal", "korean-hanja-informal", "korean-hanja-formal",
  "simp-chinese-informal", "simp-chinese-formal", "trad-chinese-informal", "trad-chinese-formal", "ethiopic-numeric",
  "thai", "devanagari", "bengali", "arabic-indic", "persian", "khmer", "lao", "myanmar", "tamil", "telugu",
  "gujarati", "gurmukhi", "kannada", "malayalam", "oriya", "tibetan", "mongolian", "cambodian", "hangul",
  "hangul-consonant",
]);

function roman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  let rest = n;
  for (const [value, letters] of table) {
    while (rest >= value) {
      out += letters;
      rest -= value;
    }
  }
  return out;
}

function alphabetic(n: number, letters: string[]): string {
  let out = "";
  let rest = n;
  while (rest > 0) {
    rest -= 1;
    out = letters[rest % letters.length]! + out;
    rest = Math.floor(rest / letters.length);
  }
  return out;
}

const LATIN = [..."abcdefghijklmnopqrstuvwxyz"];
const GREEK = [..."αβγδεζηθικλμνξοπρστυφχψω"];

/**
 * Un valor de contador escrito con un estilo de CSS, como lo dibuja el navegador. null si es un estilo que acá no se
 * calcula. Un nombre que CSS no conoce se dibuja en decimal (igual que el navegador).
 */
export function formatCounter(n: number, style: string): string | null {
  const name = style.trim().toLowerCase();
  if (name === "none") return "";
  if (name in BULLETS) return BULLETS[name]!;
  if (name === "decimal") return String(n);
  if (name === "decimal-leading-zero") return n < 0 ? `-${String(-n).padStart(2, "0")}` : String(n).padStart(2, "0");
  if (name === "lower-roman" || name === "upper-roman") {
    if (n < 1 || n > 3999) return String(n);
    const text = roman(n);
    return name === "lower-roman" ? text.toLowerCase() : text;
  }
  if (["lower-alpha", "lower-latin", "upper-alpha", "upper-latin"].includes(name)) {
    if (n < 1) return String(n);
    const text = alphabetic(n, LATIN);
    return name.startsWith("upper") ? text.toUpperCase() : text;
  }
  if (name === "lower-greek") return n < 1 ? String(n) : alphabetic(n, GREEK);
  if (UNSUPPORTED_STYLES.has(name)) return null;
  return String(n);
}

/** El texto de un ::marker (list-style-type): viñeta y espacio, o número con «. ». */
export function markerText(n: number, style: string): string | null {
  const name = style.trim().toLowerCase();
  if (name === "none") return "";
  if (name in BULLETS) return `${BULLETS[name]} `;
  const text = formatCounter(n, name);
  return text === null ? null : `${text}. `;
}

export function contentText(parts: ContentPart[], counters: ReadonlyMap<string, number>): string | null {
  let out = "";
  for (const part of parts) {
    if (part.kind === "text") {
      out += part.text;
      continue;
    }
    const text = formatCounter(counters.get(part.name) ?? 0, part.style);
    if (text === null) return null;
    out += text;
  }
  return out;
}

function upper(text: string, locale?: string): string {
  try {
    return text.toLocaleUpperCase(locale || undefined);
  } catch {
    return text.toUpperCase();
  }
}

function lower(text: string, locale?: string): string {
  try {
    return text.toLocaleLowerCase(locale || undefined);
  } catch {
    return text.toLowerCase();
  }
}

/**
 * text-transform escrito de verdad. `before` es el texto que viene antes en el mismo párrafo (para saber si una palabra
 * empieza en este nodo o viene del anterior).
 */
export function transformCase(
  text: string,
  mode: "uppercase" | "lowercase" | "capitalize",
  before = "",
  locale?: string,
): string {
  if (mode === "uppercase") return upper(text, locale);
  if (mode === "lowercase") return lower(text, locale);
  const context = before.slice(-24);
  const whole = context + text;
  let out = "";
  const segmenter = new Intl.Segmenter(locale || undefined, { granularity: "word" });
  for (const segment of segmenter.segment(whole)) {
    const end = segment.index + segment.segment.length;
    if (end <= context.length) continue;
    const piece = segment.index < context.length ? segment.segment.slice(context.length - segment.index) : segment.segment;
    if (segment.isWordLike && segment.index >= context.length) {
      const first = [...piece][0] ?? "";
      out += upper(first, locale) + piece.slice(first.length);
    } else out += piece;
  }
  return out;
}

/**
 * Versalitas: las minúsculas pasan a mayúsculas chicas (`small`), las mayúsculas y lo que no es letra quedan igual.
 * Con `all` (all-small-caps) también las mayúsculas van chicas. Los textos de `small` ya vienen en mayúscula.
 */
export function splitSmallCaps(text: string, all = false, locale?: string): Array<{ text: string; small: boolean }> {
  const out: Array<{ text: string; small: boolean }> = [];
  for (const char of text) {
    const big = upper(char, locale);
    const small = big !== char || (all && LETTER.test(char) && lower(char, locale) !== char);
    const value = small ? big : char;
    const last = out[out.length - 1];
    if (last && last.small === small) last.text += value;
    else out.push({ text: value, small });
  }
  return out;
}

/** Palabras y espacios en trozos separados (el espacio va aparte: jsafa no dibuja los nodos que son solo espacio). */
export function splitWords(text: string): string[] {
  return text.split(/([\s\u00A0]+)/).filter((piece) => piece !== "");
}

/** Los caracteres invisibles (guion opcional, espacio de ancho cero…) cada uno en su trozo. */
export function splitInvisible(text: string): string[] {
  return text.split(INVISIBLE_SPLIT).filter((piece) => piece !== "");
}

export type Crop = { left: number; top: number; right: number; bottom: number };

function percent(value: string): number | null {
  const match = /^(-?[\d.]+)%$/.exec(value.trim());
  if (match) return Number(match[1]) / 100;
  return /^0(?:px)?$/.test(value.trim()) ? 0 : null;
}

/**
 * El recorte del Word (srcRect) tal como lo deja el visor en el clip-path de la imagen: rect(arriba derecha abajo
 * izquierda) medidos desde arriba y la izquierda, o inset(...) medidos desde cada borde. null si no hay recorte.
 */
export function parseCrop(clipPath: string): Crop | null {
  const match = /^(rect|inset)\(([^)]*)\)$/.exec(clipPath.trim());
  if (!match) return null;
  const values = match[2]!.split(/\s+round\s+/)[0]!.trim().split(/\s+/).map(percent);
  if (values.length < 1 || values.length > 4 || values.some((value) => value === null)) return null;
  const [a, b = a, c = a, d = b] = values as number[];
  const crop =
    match[1] === "rect"
      ? { top: a!, right: 1 - b!, bottom: 1 - c!, left: d! }
      : { top: a!, right: b!, bottom: c!, left: d! };
  if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) return null;
  if ([crop.left, crop.top, crop.right, crop.bottom].every((value) => Math.abs(value) < 1e-6)) return null;
  return crop;
}

/**
 * Escala y corrimiento para que la parte visible del recorte ocupe justo el lugar de la imagen en la página (el visor
 * escala desde el centro, así que con recortes distintos a cada lado la imagen queda corrida).
 */
export function cropMatrix(crop: Crop, width: number, height: number) {
  const sx = 1 / (1 - crop.left - crop.right);
  const sy = 1 / (1 - crop.top - crop.bottom);
  return {
    sx,
    sy,
    tx: (sx - 1) * (width / 2) - sx * crop.left * width,
    ty: (sy - 1) * (height / 2) - sy * crop.top * height,
  };
}

const round = (value: number) => Number(value.toFixed(4));

export function cropTransform(crop: Crop, width: number, height: number, rotation = 0): string {
  const { sx, sy, tx, ty } = cropMatrix(crop, width, height);
  const turn = rotation ? `rotate(${round(rotation)}deg) ` : "";
  return `${turn}translate(${round(tx)}px, ${round(ty)}px) scale(${round(sx)}, ${round(sy)})`;
}

/** «la página 3» o «las páginas 3, 5 y 8». */
export function pagesText(pages: Iterable<number>): string {
  const list = [...new Set(pages)].filter((page) => page > 0).sort((a, b) => a - b);
  if (list.length === 0) return "algunas páginas";
  if (list.length === 1) return `la página ${list[0]}`;
  const shown = list.slice(0, 12);
  if (list.length > 12) return `las páginas ${shown.join(", ")}…`;
  return `las páginas ${shown.slice(0, -1).join(", ")} y ${shown[shown.length - 1]}`;
}

// ----- Deshacer -----

class Undo {
  private steps: Array<() => void> = [];

  push(step: () => void) {
    this.steps.push(step);
  }

  /** Pasa a esta lista lo hecho en otra (para deshacer por partes lo que no salió). */
  adopt(other: Undo) {
    this.steps.push(...other.steps);
    other.steps = [];
  }

  run() {
    for (let i = this.steps.length - 1; i >= 0; i -= 1) {
      try {
        this.steps[i]!();
      } catch {
        // Si un paso no se puede deshacer, los demás igual se deshacen.
      }
    }
    this.steps = [];
  }

  style(el: HTMLElement, prop: string, value: string, priority = "") {
    const had = el.hasAttribute("style");
    const previous = el.style.getPropertyValue(prop);
    const previousPriority = el.style.getPropertyPriority(prop);
    el.style.setProperty(prop, value, priority);
    this.push(() => {
      if (previous) el.style.setProperty(prop, previous, previousPriority);
      else el.style.removeProperty(prop);
      if (!had && el.getAttribute("style") === "") el.removeAttribute("style");
    });
  }

  addClass(el: Element, name: string) {
    if (el.classList.contains(name)) return;
    const had = el.hasAttribute("class");
    el.classList.add(name);
    this.push(() => {
      el.classList.remove(name);
      if (!had && el.getAttribute("class") === "") el.removeAttribute("class");
    });
  }

  data(el: HTMLElement, key: string, value: string) {
    const previous = el.dataset[key];
    el.dataset[key] = value;
    this.push(() => {
      if (previous === undefined) delete el.dataset[key];
      else el.dataset[key] = previous;
    });
  }

  insert(node: Node, parent: Node, before: Node | null) {
    parent.insertBefore(node, before);
    this.push(() => node.parentNode?.removeChild(node));
  }

  /** Cambia un nodo por varios; al deshacer vuelve el mismo nodo (con su texto de entonces). */
  replace(node: Node, pieces: Node[]) {
    const parent = node.parentNode;
    if (!parent || pieces.length === 0) return;
    for (const piece of pieces) parent.insertBefore(piece, node);
    parent.removeChild(node);
    this.push(() => {
      const first = pieces[0]!;
      if (first.parentNode) first.parentNode.insertBefore(node, first);
      for (const piece of pieces) piece.parentNode?.removeChild(piece);
    });
  }
}

// ----- Ayudas del DOM -----

/** Avisos agrupados por tipo, con las páginas donde pasa. */
class Notes {
  private pages = new Map<string, Set<number>>();
  private counts = new Map<string, number>();

  add(kind: string, page: number, count = 1) {
    const set = this.pages.get(kind) ?? new Set<number>();
    set.add(page);
    this.pages.set(kind, set);
    this.counts.set(kind, (this.counts.get(kind) ?? 0) + count);
  }

  flush(out: string[] | undefined) {
    if (!out) return;
    for (const [kind, set] of this.pages) {
      const where = pagesText(set);
      const count = this.counts.get(kind) ?? 0;
      out.push(MESSAGES[kind]?.(where, count) ?? `Revisá ${where} en el PDF.`);
    }
  }
}

const MESSAGES: Record<string, (where: string, count: number) => string> = {
  marker: (where) =>
    `Algunos números o viñetas de lista en ${where} no se pudieron pasar al PDF tal como se ven: revisalos en el PDF.`,
  imageBullet: (where) =>
    `Las viñetas con imagen (en ${where}) no salen en el PDF. Cambialas en el Word por una viñeta de texto.`,
  caps: (where) =>
    `Las versalitas o mayúsculas de estilo en ${where} no se pudieron pasar al PDF sin mover el texto: ahí salen como están escritas.`,
  spacing: (where) =>
    `El espaciado entre letras en ${where} no se pudo fijar sin mover el texto: ahí las letras salen un poco estiradas.`,
  manyLetters: () =>
    "Hay demasiado texto con espaciado entre letras: en el PDF esas letras salen un poco estiradas.",
  shy: (where) =>
    `Algunos guiones de corte en ${where} no se pudieron fijar: revisá el final de esos renglones en el PDF.`,
  autoHyphen: (where, count) =>
    `El visor cortó ${count === 1 ? "una palabra" : `${count} palabras`} con guion automático (en ${where}), y ese guion no sale en el PDF.`,
  autoHyphenWord: (where, count) =>
    `El Word corta palabras solo: el visor cortó ${count === 1 ? "una palabra" : `${count} palabras`} (en ${where}) y ese guion todavía no sale en el PDF. Revisá esos renglones.`,
  turned: (where) => `Hay texto girado con subrayado o tachado en ${where}: en el PDF sale sin la raya.`,
  justify: (where) =>
    `Algunos párrafos justificados en ${where} no se pudieron fijar sin mover el texto: ahí las letras pueden salir un poco estiradas.`,
  step: (where) => `Una parte de la preparación del PDF falló en ${where}. Revisá el PDF antes de mandarlo.`,
};

type Ctx = {
  root: HTMLElement;
  doc: Document;
  undo: Undo;
  notes: Notes;
  pageIndex: Map<Element, number>;
  autoHyphenation: boolean;
};

function pageOf(ctx: Ctx, node: Node | null): number {
  const el = node instanceof Element ? node : node?.parentElement;
  const section = el?.closest("section");
  return (section && ctx.pageIndex.get(section)) || 0;
}

function sectionOf(ctx: Ctx, node: Node): HTMLElement {
  const el = node instanceof Element ? node : node.parentElement;
  return (el?.closest("section.docx") as HTMLElement | null) ?? ctx.root;
}

/** Cuántos píxeles de pantalla mide un píxel CSS de la hoja (1 salvo que la hoja esté escalada). */
function scaleOf(section: HTMLElement): number {
  const width = section.getBoundingClientRect().width;
  return section.offsetWidth > 0 && width > 0 ? width / section.offsetWidth : 1;
}

function textNodes(root: Node, skip?: (node: Text) => boolean): Text[] {
  const doc = root.ownerDocument ?? (root as Document);
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const out: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text;
    if (text.data && !(skip && skip(text))) out.push(text);
  }
  return out;
}

const OURS = ".pliego-x-trazos, .pliego-x-puntos";

function charRect(node: Text, index: number): DOMRect | null {
  const code = node.data.codePointAt(index);
  if (code === undefined) return null;
  const range = node.ownerDocument.createRange();
  range.setStart(node, index);
  range.setEnd(node, Math.min(node.data.length, index + (code > 0xffff ? 2 : 1)));
  const rect = [...range.getClientRects()].find((item) => item.width > 0 && item.height > 0) ?? null;
  range.detach();
  return rect;
}

function nodeRects(node: Text): DOMRect[] {
  const range = node.ownerDocument.createRange();
  range.selectNodeContents(node);
  const rects = [...range.getClientRects()].filter((item) => item.width > 0 && item.height > 0);
  range.detach();
  return rects;
}

/** Mismo renglón: se pisan en vertical al menos la mitad del más bajo (como agrupa jsafa). */
function sameLine(a: { top: number; bottom: number }, b: { top: number; bottom: number }): boolean {
  return Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) >= 0.5 * Math.min(a.bottom - a.top, b.bottom - b.top);
}

const ratios = new Map<string, number>();

/** Dónde cae la línea de base dentro del alto de un texto con esta letra (la misma cuenta que hace jsafa). */
function baselineRatio(doc: Document, cs: CSSStyleDeclaration): number {
  const key = [cs.fontFamily, cs.fontSize, cs.fontWeight, cs.fontStyle].join("|");
  const known = ratios.get(key);
  if (known !== undefined) return known;
  const box = doc.createElement("div");
  box.style.cssText = "position:absolute;top:-99999px;left:-99999px;visibility:hidden;white-space:nowrap;";
  box.style.fontFamily = cs.fontFamily;
  box.style.fontSize = cs.fontSize;
  box.style.fontWeight = cs.fontWeight;
  box.style.fontStyle = cs.fontStyle;
  const text = doc.createTextNode("x");
  const probe = doc.createElement("span");
  probe.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline;";
  box.append(text, probe);
  doc.body.appendChild(box);
  const range = doc.createRange();
  range.selectNodeContents(text);
  const rect = range.getBoundingClientRect();
  const ratio = rect.height > 0 ? (probe.getBoundingClientRect().top - rect.top) / rect.height : 0.8;
  range.detach();
  box.remove();
  ratios.set(key, ratio);
  return ratio;
}

function blockOf(node: Node): Element | null {
  const el = node instanceof Element ? node : node.parentElement;
  return el?.closest(BLOCK) ?? el?.closest("section") ?? el ?? null;
}

function heights(blocks: Iterable<Element>): Map<Element, number> {
  const out = new Map<Element, number>();
  for (const block of blocks) out.set(block, block.getBoundingClientRect().height);
  return out;
}

function langOf(el: Element): string | undefined {
  return el.closest("[lang]")?.getAttribute("lang") || undefined;
}

/** Corre un paso; si falla, avisa y sigue con los demás (lo que alcanzó a hacer igual se deshace al final). */
function step(ctx: Ctx, fn: () => void) {
  try {
    fn();
  } catch (err) {
    console.warn("pliego: preparación del PDF", err);
    ctx.notes.add("step", 0);
  }
}

// ----- 1. Imágenes recortadas -----

const CROP_TRANSFORM = /^(?:rotate\((-?[\d.e+-]+)deg\)\s*)?scale\((-?[\d.e+-]+)(?:,\s*(-?[\d.e+-]+))?\)$/;

function fixCrops(root: HTMLElement, undo: Undo | null): number {
  let fixed = 0;
  for (const img of root.querySelectorAll("img")) {
    if (img.dataset.pliegoRecorte) continue;
    const crop = parseCrop(img.style.getPropertyValue("clip-path"));
    if (!crop) continue;
    const match = CROP_TRANSFORM.exec(img.style.transform.trim());
    const width = img.offsetWidth;
    const height = img.offsetHeight;
    if (!match || width <= 0 || height <= 0) continue;
    const transform = cropTransform(crop, width, height, Number(match[1] ?? 0));
    if (undo) {
      undo.style(img, "transform", transform);
      undo.data(img, "pliegoRecorte", "1");
    } else {
      img.style.transform = transform;
      img.dataset.pliegoRecorte = "1";
    }
    fixed += 1;
  }
  return fixed;
}

/**
 * Corrige dónde cae una imagen recortada en el Word: el visor la agranda desde el centro y, si el recorte no es igual a
 * los dos lados, la parte visible queda corrida. Llamarla después de dibujar (cambia la pantalla, no el Word).
 */
export function fixCroppedImages(root: HTMLElement): number {
  return fixCrops(root, null);
}

type Box = { x: number; y: number; width: number; height: number };
type CapturedOp = { kind: string; element?: Element; image?: { draw: Box; clip: unknown } };

/**
 * jsafa-pdf no respeta el clip-path: una imagen recortada sale entera (y agrandada) encima del texto. Sobre lo que
 * devuelve captureElement() pone el recorte justo en la parte visible. Devuelve cuántas imágenes recortó.
 */
export function clipCroppedImages(captured: { ops: ReadonlyArray<CapturedOp> }): number {
  let count = 0;
  for (const op of captured.ops) {
    if (op.kind !== "image" || !op.image || !op.element || op.element.tagName !== "IMG") continue;
    const crop = parseCrop((op.element as HTMLElement).style.getPropertyValue("clip-path"));
    if (!crop) continue;
    const { draw } = op.image;
    const x = draw.x + crop.left * draw.width;
    const y = draw.y + crop.top * draw.height;
    const width = draw.width * (1 - crop.left - crop.right);
    const height = draw.height * (1 - crop.top - crop.bottom);
    op.image.clip = [
      { type: "M", x, y },
      { type: "L", x: x + width, y },
      { type: "L", x: x + width, y: y + height },
      { type: "L", x, y: y + height },
    ];
    count += 1;
  }
  return count;
}

/** Las imágenes recortadas en el Word (para avisar si el PDF no les aplica clipCroppedImages). */
export function croppedImages(root: HTMLElement): HTMLImageElement[] {
  return [...root.querySelectorAll("img")].filter((img) => parseCrop(img.style.getPropertyValue("clip-path")));
}

// ----- 2. Números y viñetas de lista -----

const NO_COPY = new Set(["content", "counter-increment", "counter-reset", "counter-set"]);
const MARKER_PROPS = [
  "font-family", "font-size", "font-weight", "font-style", "font-stretch", "font-variant-caps", "font-variant-numeric",
  "font-feature-settings", "color", "letter-spacing", "word-spacing", "text-transform", "direction",
];

function snapshot(cs: CSSStyleDeclaration, only?: string[]): Array<[string, string]> {
  const names = only ?? Array.from({ length: cs.length }, (_, i) => cs.item(i));
  const out: Array<[string, string]> = [];
  for (const name of names) {
    if (!name || name.startsWith("--") || NO_COPY.has(name)) continue;
    out.push([name, cs.getPropertyValue(name)]);
  }
  return out;
}

/** Deja en el span, escrito en su estilo, todo lo que el pseudo-elemento tenía distinto. */
function copyStyle(span: HTMLElement, wanted: Array<[string, string]>) {
  const cs = getComputedStyle(span);
  const diff = wanted.filter(([name, value]) => cs.getPropertyValue(name) !== value);
  for (const [name, value] of diff) span.style.setProperty(name, value);
}

type Anchor = { x: number; y: number; h: number };

/** Dónde queda el primer carácter del texto (sin contar las marcas puestas acá) y el alto del párrafo. */
function anchorOf(el: Element): Anchor {
  const h = el.getBoundingClientRect().height;
  for (const node of textNodes(el)) {
    if (node.parentElement?.closest(".pliego-x-marca")) continue;
    for (let i = node.data.search(/\S/), tries = 0; i >= 0 && i < node.data.length && tries < 40; i += 1, tries += 1) {
      const rect = charRect(node, i);
      if (rect) return { x: rect.left, y: rect.top, h };
    }
  }
  return { x: Number.NaN, y: Number.NaN, h };
}

function sameAnchor(a: Anchor, b: Anchor): boolean {
  if (Math.abs(a.h - b.h) > 0.6) return false;
  if (Number.isNaN(a.x) || Number.isNaN(b.x)) return Number.isNaN(a.x) === Number.isNaN(b.x);
  return Math.abs(a.x - b.x) <= 0.6 && Math.abs(a.y - b.y) <= 0.6;
}

/** Dónde empieza el primer renglón de una lista: el primer bloque sin bloques adentro, antes de cualquier texto. */
function firstLineHost(el: Element): Element {
  let host = el;
  for (let depth = 0; depth < 8; depth += 1) {
    let next: Element | null = null;
    for (const child of host.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        if ((child as Text).data.trim()) return host;
        continue;
      }
      if (!(child instanceof Element)) continue;
      const cs = getComputedStyle(child);
      if (cs.display === "none" || cs.position === "absolute" || cs.position === "fixed") continue;
      if (["block", "list-item", "flow-root"].includes(cs.display)) next = child;
      break;
    }
    if (!next) return host;
    host = next;
  }
  return host;
}

type MarkerJob = {
  el: Element;
  kind: "before" | "inside" | "outside";
  text: string;
  wanted: Array<[string, string]>;
  host: Element;
  before: Anchor;
  span?: HTMLElement;
  undo: Undo;
};

function prepareMarkers(ctx: Ctx) {
  const { root, doc } = ctx;
  const counters = new Map<string, number>();
  const apply = (cs: CSSStyleDeclaration) => {
    for (const item of parseCounters(cs.getPropertyValue("counter-reset"))) counters.set(item.name, item.value ?? 0);
    for (const item of parseCounters(cs.getPropertyValue("counter-increment"))) {
      counters.set(item.name, (counters.get(item.name) ?? 0) + (item.value ?? 1));
    }
    for (const item of parseCounters(cs.getPropertyValue("counter-set"))) counters.set(item.name, item.value ?? 0);
  };
  // El visor reinicia los contadores en el envoltorio de las páginas (o en :root).
  const holders: Element[] = [];
  for (let el: Element | null = root; el; el = el.parentElement) holders.unshift(el);
  for (const el of root.querySelectorAll("div")) {
    if ([...el.classList].some((name) => name.endsWith("-wrapper"))) holders.push(el);
  }
  for (const el of holders) {
    for (const item of parseCounters(getComputedStyle(el).getPropertyValue("counter-reset"))) {
      counters.set(item.name, item.value ?? 0);
    }
  }

  const ordinals = new Map<Element, number>();
  const ordinalOf = (el: Element): number => {
    const owner = el.parentElement?.closest("ol, ul, menu") ?? el.parentElement ?? root;
    let current = ordinals.get(owner);
    if (current === undefined) {
      const start = owner.tagName === "OL" ? Number.parseInt(owner.getAttribute("start") ?? "1", 10) : 1;
      current = (Number.isFinite(start) ? start : 1) - 1;
    }
    const own = el.tagName === "LI" ? Number.parseInt(el.getAttribute("value") ?? "", 10) : Number.NaN;
    current = Number.isFinite(own) ? own : current + 1;
    ordinals.set(owner, current);
    return current;
  };

  const jobs: MarkerJob[] = [];
  for (const el of root.querySelectorAll("p, li")) {
    if (el.closest(OURS)) continue;
    const cs = getComputedStyle(el);
    if (cs.display === "none") continue;
    apply(cs);
    const pseudo = getComputedStyle(el, "::before");
    const content = pseudo.getPropertyValue("content");
    if (content && content !== "none" && content !== "normal") {
      apply(pseudo);
      const parts = parseContent(content);
      const text = parts ? contentText(parts, counters) : null;
      if (text !== null && !text.trim() && pseudo.backgroundImage !== "none") ctx.notes.add("imageBullet", pageOf(ctx, el));
      else if (text === null) ctx.notes.add("marker", pageOf(ctx, el));
      else if (text) {
        const wanted = snapshot(pseudo);
        jobs.push({ el, kind: "before", text, wanted, host: el, before: { x: 0, y: 0, h: 0 }, undo: new Undo() });
      }
    }
    if (cs.display !== "list-item") continue;
    const n = ordinalOf(el);
    if (cs.listStyleType === "none" && cs.listStyleImage === "none") continue;
    if (cs.listStyleImage !== "none") {
      ctx.notes.add("imageBullet", pageOf(ctx, el));
      continue;
    }
    const text = markerText(n, cs.listStyleType);
    if (text === null) {
      ctx.notes.add("marker", pageOf(ctx, el));
      continue;
    }
    if (!text) continue;
    const marker = getComputedStyle(el, "::marker");
    const wanted = snapshot(marker, MARKER_PROPS);
    wanted.push(["white-space", "pre"]);
    jobs.push({
      el,
      kind: cs.listStylePosition === "inside" ? "inside" : "outside",
      text,
      wanted,
      host: firstLineHost(el),
      before: { x: 0, y: 0, h: 0 },
      undo: new Undo(),
    });
  }
  if (jobs.length === 0) return;

  for (const job of jobs) job.before = anchorOf(job.el);
  for (const job of jobs) {
    const span = doc.createElement("span");
    span.className = "pliego-x-marca";
    span.textContent = job.text;
    job.undo.insert(span, job.host, job.host.firstChild);
    job.undo.addClass(job.el, job.kind === "before" ? "pliego-x-sin-antes" : "pliego-x-sin-vineta");
    job.span = span;
  }
  for (const job of jobs) if (job.span) copyStyle(job.span, job.wanted);
  // La marca de afuera queda a la izquierda del renglón: ocupa su ancho y lo devuelve con un margen negativo.
  const widths = jobs.map((job) =>
    job.kind === "outside" && job.span ? job.span.getBoundingClientRect().width / scaleOf(sectionOf(ctx, job.el)) : 0,
  );
  jobs.forEach((job, i) => {
    if (job.kind === "outside" && job.span) job.span.style.marginLeft = `${-(widths[i] ?? 0)}px`;
  });
  for (const job of jobs) {
    const after = anchorOf(job.el);
    if (sameAnchor(job.before, after)) ctx.undo.adopt(job.undo);
    else {
      job.undo.run();
      ctx.notes.add("marker", pageOf(ctx, job.el));
    }
  }
}

// ----- 3. Mayúsculas de estilo y versalitas -----

type CaseJob = { node: Text; transform: "uppercase" | "lowercase" | "capitalize" | null; caps: "small" | "all" | null; before: string };

function prepareCase(ctx: Ctx) {
  const { doc } = ctx;
  const styles = new Map<Element, { transform: CaseJob["transform"]; caps: CaseJob["caps"] }>();
  const jobs: CaseJob[] = [];
  let lastBlock: Element | null = null;
  let lastText = "";
  for (const node of textNodes(ctx.root, (text) => Boolean(text.parentElement?.closest(OURS)))) {
    const parent = node.parentElement;
    if (!parent) continue;
    let style = styles.get(parent);
    if (!style) {
      const cs = getComputedStyle(parent);
      const words = cs.textTransform.split(/\s+/);
      const transform = (["uppercase", "lowercase", "capitalize"] as const).find((mode) => words.includes(mode)) ?? null;
      const variant = cs.fontVariantCaps;
      const caps = variant === "all-small-caps" || variant === "all-petite-caps"
        ? "all"
        : variant === "small-caps" || variant === "petite-caps"
          ? "small"
          : null;
      style = { transform, caps };
      styles.set(parent, style);
    }
    const block = blockOf(node);
    if (block !== lastBlock) lastText = "";
    lastBlock = block;
    if (style.transform || style.caps) jobs.push({ node, ...style, before: lastText });
    lastText = (lastText + node.data).slice(-24);
  }
  if (jobs.length === 0) return;

  const byBlock = new Map<Element, CaseJob[]>();
  for (const job of jobs) {
    const block = blockOf(job.node) ?? ctx.root;
    byBlock.set(block, [...(byBlock.get(block) ?? []), job]);
  }
  const before = heights(byBlock.keys());
  const undos = new Map<Element, Undo>();
  for (const [block, list] of byBlock) {
    const undo = new Undo();
    undos.set(block, undo);
    for (const job of list) {
      const parent = job.node.parentElement;
      if (!parent) continue;
      const locale = langOf(parent);
      let text = job.node.data;
      if (job.transform) {
        text = transformCase(text, job.transform, job.before, locale);
        undo.addClass(parent, "pliego-x-sin-transformar");
      }
      if (job.caps) {
        undo.addClass(parent, "pliego-x-sin-versalitas");
        const pieces = splitSmallCaps(text, job.caps === "all", locale).map((piece) => {
          if (!piece.small) return doc.createTextNode(piece.text);
          const span = doc.createElement("span");
          span.className = "pliego-x-versalita";
          // El navegador arma las versalitas que la letra no trae con mayúsculas al 70 %.
          span.style.fontSize = "0.7em";
          span.textContent = piece.text;
          return span;
        });
        undo.replace(job.node, pieces);
      } else if (text !== job.node.data) undo.push(setText(job.node, text));
    }
  }
  const after = heights(byBlock.keys());
  for (const [block, undo] of undos) {
    if (Math.abs((after.get(block) ?? 0) - (before.get(block) ?? 0)) <= 0.6) ctx.undo.adopt(undo);
    else {
      undo.run();
      ctx.notes.add("caps", pageOf(ctx, block));
    }
  }
}

function setText(node: Text, value: string): () => void {
  const previous = node.data;
  node.data = value;
  return () => {
    node.data = previous;
  };
}

// ----- 4. Guiones opcionales y caracteres invisibles -----

/** Arma los trozos de un nodo: cada invisible aparte y, donde el renglón se corta en un guion opcional, un guion real. */
function invisiblePieces(doc: Document, data: string, broken: Set<number>) {
  const pieces: Text[] = [];
  const hyphens: Array<{ dash: Text; zwsp: Text }> = [];
  let buffer = "";
  let index = 0;
  const flush = () => {
    if (buffer) pieces.push(doc.createTextNode(buffer));
    buffer = "";
  };
  for (const piece of splitInvisible(data)) {
    if (piece === SHY && broken.has(index)) {
      buffer += "-";
      flush();
      const dash = pieces[pieces.length - 1]!;
      const zwsp = doc.createTextNode(ZWSP);
      pieces.push(zwsp);
      hyphens.push({ dash, zwsp });
    } else if (INVISIBLE.test(piece) && piece.length === 1) {
      flush();
      pieces.push(doc.createTextNode(piece));
    } else buffer += piece;
    index += piece.length;
  }
  flush();
  return { pieces, hyphens };
}

/** El primer carácter visible después de un nodo, dentro del mismo párrafo. */
function nextVisible(ctx: Ctx, from: Text): DOMRect | null {
  const block = blockOf(from) ?? ctx.root;
  const walker = ctx.doc.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  walker.currentNode = from;
  for (let node = walker.nextNode(), seen = 0; node && seen < 20; node = walker.nextNode(), seen += 1) {
    const text = node as Text;
    for (let i = 0; i < text.data.length && i < 8; i += 1) {
      const rect = charRect(text, i);
      if (rect) return rect;
    }
  }
  return null;
}

function prepareHyphens(ctx: Ctx) {
  const { doc } = ctx;
  const found = textNodes(ctx.root, (text) => !INVISIBLE.test(text.data) || Boolean(text.parentElement?.closest(OURS)));
  if (found.length === 0) return;
  // Un guion opcional donde se corta el renglón se dibuja como guion: tiene ancho. Donde no se corta, mide cero.
  const breaks = found.map((node) => {
    const broken = new Set<number>();
    for (let i = node.data.indexOf(SHY); i >= 0; i = node.data.indexOf(SHY, i + 1)) {
      const range = doc.createRange();
      range.setStart(node, i);
      range.setEnd(node, i + 1);
      if ([...range.getClientRects()].some((rect) => rect.width > 0.5)) broken.add(i);
      range.detach();
    }
    return broken;
  });
  const jobs = found.map((node, i) => {
    const undo = new Undo();
    const { pieces, hyphens } = invisiblePieces(doc, node.data, breaks[i]!);
    undo.replace(node, pieces);
    return { node, undo, hyphens };
  });
  for (const job of jobs) {
    const ok = job.hyphens.every(({ dash, zwsp }) => {
      const end = charRect(dash, dash.data.length - 1);
      const next = nextVisible(ctx, zwsp);
      return Boolean(end && (!next || !sameLine(end, next)));
    });
    if (ok) {
      ctx.undo.adopt(job.undo);
      continue;
    }
    // Si el guion real mueve el corte, vuelve el opcional (al menos fuera de la palabra en el PDF) y se avisa.
    job.undo.run();
    ctx.notes.add("shy", pageOf(ctx, job.node));
    const { pieces } = invisiblePieces(doc, job.node.data, new Set());
    ctx.undo.replace(job.node, pieces);
  }
}

/** Cuenta palabras que el navegador cortó solo (hyphens: auto): ese guion no está en el texto y no sale en el PDF. */
function countAutoHyphens(ctx: Ctx) {
  const hyphenating = new Map<Element, boolean>();
  for (const node of textNodes(ctx.root, (text) => Boolean(text.parentElement?.closest(OURS)))) {
    const parent = node.parentElement;
    if (!parent || node.data.length < 4) continue;
    let auto = hyphenating.get(parent);
    if (auto === undefined) {
      auto = getComputedStyle(parent).hyphens === "auto";
      hyphenating.set(parent, auto);
    }
    if (!auto) continue;
    let breaks = 0;
    for (const start of lineStarts(node)) {
      if (LETTER.test(node.data[start - 1] ?? "") && LETTER.test(node.data[start] ?? "")) breaks += 1;
    }
    if (breaks) ctx.notes.add(ctx.autoHyphenation ? "autoHyphenWord" : "autoHyphen", pageOf(ctx, node), breaks);
  }
}

/** Dónde empieza cada renglón dentro de un nodo de texto (sin el primero). Busca a saltos, sin medir letra por letra. */
function lineStarts(node: Text): number[] {
  const lines: Array<{ top: number; bottom: number }> = [];
  for (const rect of nodeRects(node)) {
    if (!lines.some((line) => sameLine(line, rect))) lines.push({ top: rect.top, bottom: rect.bottom });
  }
  if (lines.length < 2) return [];
  lines.sort((a, b) => a.top - b.top);
  const length = node.data.length;
  // Renglón del primer carácter visible desde i (los espacios del final de un renglón no tienen caja).
  const lineAt = (i: number): number => {
    for (let j = i, tries = 0; j < length && tries < 64; j += 1, tries += 1) {
      const rect = charRect(node, j);
      if (!rect) continue;
      const index = lines.findIndex((line) => sameLine(line, rect));
      return index < 0 ? lines.length : index;
    }
    return lines.length;
  };
  const starts: number[] = [];
  for (let k = 1; k < lines.length; k += 1) {
    let lo = starts[starts.length - 1] ?? 0;
    let hi = length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (lineAt(mid) >= k) hi = mid;
      else lo = mid + 1;
    }
    if (lo > 0 && lo < length) starts.push(lo);
  }
  return starts;
}

// ----- 5. Puntos guía de los tabuladores -----

function prepareTabLeaders(ctx: Ctx) {
  const { doc } = ctx;
  const tabs = [...ctx.root.querySelectorAll("span")].filter((span) =>
    [...span.classList].some((name) => name.endsWith("-tab-stop")),
  );
  const dotted = tabs.filter((tab) => {
    const cs = getComputedStyle(tab);
    return cs.textDecorationLine.includes("underline") && cs.textDecorationStyle === "dotted";
  });
  const sizes = dotted.map((tab) => {
    const rect = tab.getBoundingClientRect();
    const scale = scaleOf(sectionOf(ctx, tab));
    return { width: rect.width / scale, height: rect.height / scale };
  });
  dotted.forEach((tab, i) => {
    const { width, height } = sizes[i]!;
    if (width < 4 || height <= 0) return;
    // Word pone puntos sobre la línea de base; la pantalla los muestra como subrayado punteado (que jsafa no dibuja).
    ctx.undo.style(tab, "position", "relative");
    ctx.undo.style(tab, "text-decoration-line", "none");
    const dots = doc.createElement("span");
    dots.className = "pliego-x-puntos";
    dots.style.cssText =
      `position:absolute;left:0;top:0;width:${width}px;height:${height}px;line-height:${height}px;` +
      "white-space:pre;text-align:right;text-indent:0;text-decoration:none;word-spacing:0;letter-spacing:0;overflow:visible;";
    dots.textContent = ".".repeat(20);
    ctx.undo.insert(dots, tab, null);
  });
  for (const dots of ctx.root.querySelectorAll<HTMLElement>(".pliego-x-puntos")) {
    const text = dots.firstChild;
    if (!(text instanceof Text) || text.data !== ".".repeat(20)) continue;
    const rects = nodeRects(text);
    const scale = scaleOf(sectionOf(ctx, dots));
    const advance = (rects[0]?.width ?? 0) / scale / 20;
    const width = Number.parseFloat(dots.style.width);
    const count = advance > 0 ? Math.floor((width - advance * 0.5) / advance) : 0;
    if (count < 1) dots.textContent = "";
    else text.data = ".".repeat(count);
  }
}

// ----- 6. Subrayado, tachado y sobrerrayado -----

type Deco = { lines: string[]; style: string; color: string; thickness: number; size: number };
type Info = { deco: Deco | null; stop: boolean; turned: boolean; section: boolean };

function infoOf(cache: Map<Element, Info>, el: Element): Info {
  const known = cache.get(el);
  if (known) return known;
  const cs = getComputedStyle(el);
  const lines = cs.textDecorationLine.split(/\s+/).filter((line) => ["underline", "line-through", "overline"].includes(line));
  const size = Number.parseFloat(cs.fontSize) || 16;
  const thickness = Number.parseFloat(cs.textDecorationThickness);
  const color = cs.textDecorationColor;
  const deco =
    lines.length && !/rgba\(\s*0,\s*0,\s*0,\s*0\s*\)|transparent/.test(color)
      ? {
          lines,
          style: cs.textDecorationStyle,
          color,
          thickness: Number.isFinite(thickness) && thickness > 0 ? thickness : Math.max(0.5, size * 0.05),
          size,
        }
      : null;
  const info: Info = {
    deco,
    // La raya del padre no entra en un inline-block, un flotante ni algo posicionado aparte.
    stop: /^inline-/.test(cs.display) || cs.float !== "none" || cs.position === "absolute" || cs.position === "fixed",
    turned: cs.transform !== "none",
    section: el.tagName === "SECTION",
  };
  cache.set(el, info);
  return info;
}

type Piece = { left: number; right: number; top: number; bottom: number; baseline: number; aligned: boolean };

function prepareDecorations(ctx: Ctx) {
  const { doc } = ctx;
  const cache = new Map<Element, Info>();
  const buckets = new Map<Deco, Map<string, { section: HTMLElement; pieces: Piece[] }>>();
  const fonts = new Map<Element, { ratio: number; aligned: boolean }>();
  for (const node of textNodes(ctx.root, (text) => Boolean(text.parentElement?.closest(OURS)))) {
    const parent = node.parentElement;
    if (!parent) continue;
    const found = new Map<string, Deco>();
    let turned = false;
    for (let el: Element | null = parent; el; el = el.parentElement) {
      const info = infoOf(cache, el);
      if (info.turned) turned = true;
      if (info.deco) for (const line of info.deco.lines) if (!found.has(line)) found.set(line, info.deco);
      if (info.stop || info.section || el === ctx.root) break;
    }
    if (found.size === 0) continue;
    if (turned) {
      ctx.notes.add("turned", pageOf(ctx, node));
      continue;
    }
    let font = fonts.get(parent);
    if (!font) {
      const cs = getComputedStyle(parent);
      font = { ratio: baselineRatio(doc, cs), aligned: cs.verticalAlign === "baseline" };
      fonts.set(parent, font);
    }
    const rects = nodeRects(node);
    if (rects.length === 0) continue;
    const section = sectionOf(ctx, node);
    for (const [line, deco] of found) {
      const lines = buckets.get(deco) ?? new Map<string, { section: HTMLElement; pieces: Piece[] }>();
      buckets.set(deco, lines);
      const bucket = lines.get(line) ?? { section, pieces: [] };
      lines.set(line, bucket);
      for (const rect of rects) {
        bucket.pieces.push({
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          baseline: rect.top + font.ratio * rect.height,
          aligned: font.aligned,
        });
      }
    }
  }
  if (buckets.size === 0) return;

  const layers = new Map<HTMLElement, { layer: HTMLElement; scale: number; left: number; top: number }>();
  const layerOf = (section: HTMLElement) => {
    const known = layers.get(section);
    if (known) return known;
    const layer = doc.createElement("div");
    layer.className = "pliego-x-trazos";
    layer.setAttribute("aria-hidden", "true");
    layer.style.cssText = "position:absolute;left:0;top:0;width:0;height:0;overflow:visible;pointer-events:none;z-index:2;";
    const rect = section.getBoundingClientRect();
    const scale = scaleOf(section);
    const entry = { layer, scale, left: rect.left + section.clientLeft * scale, top: rect.top + section.clientTop * scale };
    layers.set(section, entry);
    return entry;
  };

  const drawn: Array<{ section: HTMLElement; boxes: string[] }> = [];
  for (const [deco, lines] of buckets) {
    for (const [line, bucket] of lines) {
      const { scale, left: originX, top: originY } = layerOf(bucket.section);
      const boxes: string[] = [];
      for (const run of joinRuns(bucket.pieces)) {
        const baselines = run.pieces.filter((piece) => piece.aligned).map((piece) => piece.baseline);
        const baseline = Math.max(...(baselines.length ? baselines : run.pieces.map((piece) => piece.baseline)));
        const th = deco.thickness;
        const size = deco.size;
        const y =
          line === "underline"
            ? (baseline - originY) / scale + size * 0.08
            : line === "line-through"
              ? (baseline - originY) / scale - size * 0.3 - th / 2
              : (run.top - originY) / scale;
        const x = (run.left - originX) / scale;
        const width = (run.right - run.left) / scale;
        boxes.push(...rule(x, width, y, th, deco.style, deco.color, line === "underline" ? 1 : line === "overline" ? -1 : 0));
      }
      drawn.push({ section: bucket.section, boxes });
    }
  }
  for (const { section, boxes } of drawn) {
    if (boxes.length === 0) continue;
    const { layer } = layerOf(section);
    if (!layer.parentNode) ctx.undo.insert(layer, section, null);
    layer.insertAdjacentHTML("beforeend", boxes.join(""));
  }
}

/** Junta los trozos de un mismo renglón que se tocan (un hueco, como una imagen en el medio, corta la raya). */
function joinRuns(pieces: Piece[]) {
  const sorted = [...pieces].sort((a, b) => a.top - b.top || a.left - b.left);
  const lines: Piece[][] = [];
  for (const piece of sorted) {
    const line = lines.find((items) => sameLine(items[0]!, piece));
    if (line) line.push(piece);
    else lines.push([piece]);
  }
  const runs: Array<{ left: number; right: number; top: number; pieces: Piece[] }> = [];
  for (const line of lines) {
    line.sort((a, b) => a.left - b.left);
    let current: { left: number; right: number; top: number; pieces: Piece[] } | null = null;
    for (const piece of line) {
      if (current && piece.left <= current.right + 1.5) {
        current.right = Math.max(current.right, piece.right);
        current.top = Math.min(current.top, piece.top);
        current.pieces.push(piece);
      } else {
        current = { left: piece.left, right: piece.right, top: piece.top, pieces: [piece] };
        runs.push(current);
      }
    }
  }
  return runs;
}

/** Las cajas de una raya (jsafa dibuja el fondo de una caja; un borde punteado lo dibujaría liso). */
function rule(x: number, width: number, y: number, th: number, style: string, color: string, side: number): string[] {
  const box = (left: number, top: number, w: number, h: number) =>
    `<div style="position:absolute;left:${round(left)}px;top:${round(top)}px;width:${round(w)}px;height:${round(h)}px;background-color:${color};"></div>`;
  if (width <= 0) return [];
  if (style === "double") {
    const gap = th * 2;
    const second = side < 0 ? y - gap : y + gap;
    return [box(x, y, width, th), box(x, second, width, th)];
  }
  const dash = style === "dotted" ? th : style === "dashed" ? th * 3 : 0;
  const gap = style === "dotted" ? th : th * 2;
  if (dash > 0 && width / (dash + gap) <= 3000) {
    const out: string[] = [];
    for (let left = x; left < x + width - 0.01; left += dash + gap) out.push(box(left, y, Math.min(dash, x + width - left), th));
    return out;
  }
  // Ondulado (y lo que tenga demasiados trazos): raya lisa.
  return [box(x, y, width, th)];
}

// ----- 7. Justificado y espaciado entre letras -----

function prepareSpacing(ctx: Ctx) {
  const { doc } = ctx;
  const styles = new Map<Element, { justify: boolean; spacing: number }>();
  const words: Text[] = [];
  const letters: Array<{ node: Text; spacing: number }> = [];
  let letterCount = 0;
  for (const node of textNodes(ctx.root, (text) => !text.data.trim() || Boolean(text.parentElement?.closest(OURS)))) {
    const parent = node.parentElement;
    if (!parent) continue;
    let style = styles.get(parent);
    if (!style) {
      const cs = getComputedStyle(parent);
      const spacing = Number.parseFloat(cs.letterSpacing);
      style = {
        justify: cs.textAlign === "justify" || cs.textAlignLast === "justify",
        spacing: Number.isFinite(spacing) && Math.abs(spacing) >= 0.05 ? spacing : 0,
      };
      styles.set(parent, style);
    }
    if (style.spacing) {
      letters.push({ node, spacing: style.spacing });
      letterCount += node.data.length;
    } else if (style.justify && /[\s\u00A0]/.test(node.data)) words.push(node);
  }
  if (letterCount > MAX_SPACED_LETTERS) {
    letters.length = 0;
    ctx.notes.add("manyLetters", 0);
  }
  if (words.length === 0 && letters.length === 0) return;

  const byBlock = new Map<Element, Undo>();
  const undoFor = (node: Node) => {
    const block = blockOf(node) ?? ctx.root;
    const undo = byBlock.get(block) ?? new Undo();
    byBlock.set(block, undo);
    return undo;
  };
  for (const node of [...words, ...letters.map((item) => item.node)]) undoFor(node);
  const before = heights(byBlock.keys());

  // Cada palabra en su nodo: jsafa la pone en su lugar y no estira el renglón entero hasta el margen.
  for (const node of words) {
    undoFor(node).replace(node, splitWords(node.data).map((piece) => doc.createTextNode(piece)));
  }
  // Cada letra en su caja, con el espacio después como margen: si no, jsafa estira cada letra hasta ocupar ese espacio.
  const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
  for (const { node, spacing } of letters) {
    const pieces: Node[] = [];
    for (const { segment } of segmenter.segment(node.data)) {
      // Los espacios quedan juntos como texto (jsafa no los dibuja); un invisible va solo, como en el paso de guiones.
      if (!segment.trim() && !INVISIBLE.test(segment)) {
        const last = pieces[pieces.length - 1];
        if (last instanceof Text && !last.data.trim() && !INVISIBLE.test(last.data)) last.data += segment;
        else pieces.push(doc.createTextNode(segment));
        continue;
      }
      if (INVISIBLE.test(segment)) {
        pieces.push(doc.createTextNode(segment));
        continue;
      }
      const span = doc.createElement("span");
      span.className = "pliego-x-letra";
      span.style.letterSpacing = "0px";
      span.style.marginRight = `${spacing}px`;
      span.textContent = segment;
      pieces.push(span);
    }
    undoFor(node).replace(node, pieces);
  }

  const after = heights(byBlock.keys());
  const spaced = new Set(letters.map((item) => blockOf(item.node)));
  for (const [block, undo] of byBlock) {
    if (Math.abs((after.get(block) ?? 0) - (before.get(block) ?? 0)) <= 0.6) ctx.undo.adopt(undo);
    else {
      undo.run();
      ctx.notes.add(spaced.has(block) ? "spacing" : "justify", pageOf(ctx, block));
    }
  }
}

// ----- Todo junto -----

/**
 * Prepara las páginas ya dibujadas para que jsafa-pdf las copie como se ven. Llamarla con la clase `pliego-exporting`
 * puesta (si falta, la pone) y después de que la hoja esté a escala 1. Devuelve con qué deshacer todo, en orden.
 *
 * Con `autoHyphenation` (el Word cortaba palabras solo) no se imita ese corte: solo se avisa cuántas palabras cortó el
 * visor con guion automático, porque ese guion no está en el texto. Los guiones opcionales escritos en el Word sí se
 * fijan siempre.
 */
export function prepareForExport(root: HTMLElement, opts: ExportPrepOptions): () => void {
  const undo = new Undo();
  const ctx: Ctx = {
    root,
    doc: root.ownerDocument,
    undo,
    notes: new Notes(),
    pageIndex: new Map(),
    autoHyphenation: opts.autoHyphenation,
  };
  const pages = root.matches("section.docx") ? [root] : [...root.querySelectorAll("section.docx")];
  pages.forEach((page, i) => ctx.pageIndex.set(page, i + 1));
  if (!root.closest(".pliego-exporting")) undo.addClass(root, "pliego-exporting");

  step(ctx, () => fixCrops(root, undo));
  step(ctx, () => prepareMarkers(ctx));
  step(ctx, () => prepareCase(ctx));
  step(ctx, () => prepareHyphens(ctx));
  step(ctx, () => countAutoHyphens(ctx));
  step(ctx, () => prepareTabLeaders(ctx));
  step(ctx, () => prepareDecorations(ctx));
  step(ctx, () => prepareSpacing(ctx));
  ctx.notes.flush(opts.warnings);

  let done = false;
  return () => {
    if (done) return;
    done = true;
    undo.run();
  };
}

import { FOOTNOTE_MARK, MARK_CLOSE, MARK_OPEN, PAGES_MARK, SECTION_PAGES_MARK, type SectionNumbering } from "./fields.ts";
import { DEFAULT_RENDER_INFO, type RenderInfo, type SectionLayout } from "./render-copy.ts";

const PAGE_SELECTOR = "section.docx";
const NOTE_REF = "sup[data-pliego-note]";
/** Nota con marca propia (w:customMarkFollows): la llamada no lleva número, el símbolo ya está en el texto. */
export const CUSTOM_NOTE_MARK = `${MARK_OPEN}C${MARK_CLOSE}`;
/** Puntos a píxeles CSS. */
const PT = 96 / 72;

function pagesOf(root: HTMLElement): HTMLElement[] {
  return [...root.querySelectorAll(PAGE_SELECTOR)] as HTMLElement[];
}

function childrenOf(page: HTMLElement, tag: string): HTMLElement[] {
  return ([...page.children] as HTMLElement[]).filter((child) => child.tagName.toLowerCase() === tag);
}

/** Alto de la hoja según el Word (el visor lo pone como alto mínimo; si el contenido no entra, la página crece). */
function sheetHeight(page: HTMLElement): number {
  const value = Number.parseFloat(getComputedStyle(page).minHeight);
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function overflows(page: HTMLElement, height: number): boolean {
  return height > 0 && page.offsetHeight > height + 1;
}

function markTexts(root: HTMLElement): Text[] {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const found: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeValue?.includes(MARK_OPEN)) found.push(node as Text);
  }
  return found;
}

/** La lista de notas al pie de la página (el visor la pone después del texto, antes del pie). */
function notesList(page: HTMLElement): HTMLElement | null {
  return childrenOf(page, "ol")[0] ?? null;
}

/** Página en blanco que se agregó para que una sección empiece en página impar o par. */
function isBlank(page: HTMLElement): boolean {
  return page.dataset.pliegoBlank !== undefined;
}

export type PageMap = {
  /** Sección de cada página. */
  section: Map<HTMLElement, number>;
  /** Páginas donde empieza una sección (el visor eligió para ellas el encabezado de primera página). */
  first: Set<HTMLElement>;
  /** Encabezado y pie del Word para cada sección, según la página sea impar (0) o par (1). */
  templates: Map<number, Array<{ header: HTMLElement | null; footer: HTMLElement | null } | undefined>>;
};

/** La llamada a nota que sigue a una marca: se saltean los textos que son solo marcas (o quedaron vacíos). */
function supAfter(node: Text): HTMLElement | null {
  const onlyMarks = new RegExp(`^(?:${MARK_OPEN}[^${MARK_CLOSE}]*${MARK_CLOSE})*$`);
  let next = node.nextSibling;
  while (next && next.nodeType === Node.TEXT_NODE && onlyMarks.test(next.nodeValue ?? "")) next = next.nextSibling;
  return next instanceof HTMLElement && next.tagName === "SUP" ? next : null;
}

/**
 * Quita del texto las marcas de sección y de notas al pie, y anota de qué sección es cada página y qué encabezado y pie
 * le tocan (para repetirlos bien en las páginas que se agreguen).
 */
export function takeMarks(root: HTMLElement): PageMap {
  const pages = pagesOf(root);
  const section = new Map<HTMLElement, number>();
  const first = new Set<HTMLElement>();
  const starts = new Map<HTMLElement, number>();
  const pattern = new RegExp(`${MARK_OPEN}S(\\d+)${MARK_CLOSE}`, "g");
  const nodes = markTexts(root);
  // Primero se buscan las llamadas (con las marcas todavía puestas) y después se limpia el texto.
  for (const node of nodes) {
    const text = node.nodeValue ?? "";
    if (!text.includes(FOOTNOTE_MARK) && !text.includes(CUSTOM_NOTE_MARK)) continue;
    const sup = supAfter(node);
    if (!sup) continue;
    if (text.includes(FOOTNOTE_MARK)) sup.dataset.pliegoNote = "";
    if (text.includes(CUSTOM_NOTE_MARK)) sup.dataset.pliegoCustom = "";
  }
  for (const node of nodes) {
    let text = node.nodeValue ?? "";
    for (const match of text.matchAll(pattern)) {
      const page = pages.find((item) => item.contains(node));
      if (page && !starts.has(page)) starts.set(page, Number(match[1]));
    }
    text = text.replace(pattern, "").split(FOOTNOTE_MARK).join("").split(CUSTOM_NOTE_MARK).join("");
    node.nodeValue = text;
  }
  // Una marca de nota propia sin su marca de nota común igual es una llamada (por si llegan en textos separados).
  for (const sup of root.querySelectorAll("sup[data-pliego-custom]")) (sup as HTMLElement).dataset.pliegoNote = "";
  const templates: PageMap["templates"] = new Map();
  let current = 0;
  pages.forEach((page, index) => {
    const begins = starts.get(page);
    if (begins !== undefined) {
      current = begins;
      first.add(page);
    }
    section.set(page, current);
    if (first.has(page)) return;
    const slots = templates.get(current) ?? [];
    slots[index % 2] ??= { header: childrenOf(page, "header")[0] ?? null, footer: childrenOf(page, "footer")[0] ?? null };
    templates.set(current, slots);
  });
  return { section, first, templates };
}

/** Cambia el encabezado o el pie de una página por una copia del modelo. */
function swap(page: HTMLElement, tag: "header" | "footer", model: HTMLElement | null) {
  if (!model) return;
  const current = childrenOf(page, tag)[0];
  if (current?.isEqualNode(model)) return;
  const copy = model.cloneNode(true) as HTMLElement;
  if (current) current.replaceWith(copy);
  else if (tag === "header") page.prepend(copy);
  else page.append(copy);
}

/**
 * Pasa a páginas nuevas lo que no entra en la hoja, igual que haría Word: se mueven párrafos o tablas enteros desde
 * el final, con sus notas al pie. Se puede llamar otra vez (después de cambiar márgenes): solo toca las páginas que
 * no entran. Devuelve cuántas páginas agregó y anota las que siguen sin entrar (un solo bloque más alto que la hoja).
 */
function flow(root: HTMLElement, map: PageMap, tooTall: Set<HTMLElement>): number {
  let added = 0;
  let page = root.querySelector(PAGE_SELECTOR) as HTMLElement | null;
  while (page) {
    const current = page;
    suppressTopSpace(current);
    const height = sheetHeight(current);
    const blocks = () => childrenOf(current, "article").reduce((sum, article) => sum + article.children.length, 0);
    const lastBlock = () => {
      const articles = childrenOf(current, "article").filter((article) => article.children.length > 0);
      return (articles[articles.length - 1]?.lastElementChild as HTMLElement | null) ?? null;
    };
    const canSplit = () => {
      const block = lastBlock();
      return Boolean(block && ((block.tagName === "P" && linesOf(block).length >= 4) || (block.tagName === "TABLE" && (block as HTMLTableElement).rows.length > 1)));
    };
    if (overflows(current, height) && (blocks() > 1 || canSplit())) {
      const fresh = current.cloneNode(false) as HTMLElement;
      delete fresh.dataset.pliegoBlank;
      const header = childrenOf(current, "header")[0];
      const footer = childrenOf(current, "footer")[0];
      if (header) fresh.appendChild(header.cloneNode(true));
      if (footer) fresh.appendChild(footer.cloneNode(true));
      current.after(fresh);
      map.section.set(fresh, map.section.get(current) ?? 0);
      added += 1;
      const freshArticles = new Map<HTMLElement, HTMLElement>();
      let freshNotes: HTMLElement | null = null;
      while (overflows(current, height) && (blocks() > 1 || canSplit())) {
        const articles = childrenOf(current, "article").filter((article) => article.children.length > 0);
        const source = articles[articles.length - 1]!;
        let target = freshArticles.get(source);
        if (!target) {
          target = source.cloneNode(false) as HTMLElement;
          freshArticles.set(source, target);
          const before = childrenOf(fresh, "article")[0] ?? childrenOf(fresh, "ol")[0] ?? childrenOf(fresh, "footer")[0];
          fresh.insertBefore(target, before ?? null);
        }
        const last = source.lastElementChild as HTMLElement;
        let block: HTMLElement = last;
        if (last.tagName === "TABLE" && (last as HTMLTableElement).rows.length > 1) {
          // Una tabla que no entra se parte por filas: las que sobran van a una copia de la tabla en la página nueva.
          const table = last as HTMLTableElement;
          const tail = table.cloneNode(false) as HTMLTableElement;
          tail.removeAttribute("id");
          for (const part of [...table.children]) if (part.tagName === "COLGROUP") tail.appendChild(part.cloneNode(true));
          target.insertBefore(tail, target.firstChild);
          while (overflows(current, height) && table.rows.length > 1) {
            const row = table.rows[table.rows.length - 1]!;
            const firstRow = tail.rows[0];
            if (firstRow) firstRow.before(row);
            else tail.appendChild(row);
          }
          block = tail;
        } else {
          const tail = splitParagraph(last, current.offsetHeight - height);
          block = tail ?? last;
          target.insertBefore(block, target.firstChild);
        }
        // Las notas de las llamadas que se movieron son las últimas de la lista de esta página.
        const moved = block.querySelectorAll(NOTE_REF).length;
        const notes = notesList(current);
        if (moved && notes) {
          if (!freshNotes) {
            freshNotes = notes.cloneNode(false) as HTMLElement;
            fresh.insertBefore(freshNotes, childrenOf(fresh, "footer")[0] ?? null);
          }
          freshNotes.prepend(...[...notes.children].slice(-moved));
          if (!notes.children.length) notes.remove();
        }
        const left = childrenOf(current, "article");
        for (const article of left) if (!article.children.length && left.length > 1) article.remove();
      }
    }
    if (overflows(current, height)) tooTall.add(current);
    else tooTall.delete(current);
    const following = current.nextElementSibling as HTMLElement | null;
    page = following?.matches(PAGE_SELECTOR) ? following : null;
  }
  return added;
}

/** Los renglones de un párrafo (cajas agrupadas por altura), de arriba abajo. */
function linesOf(block: HTMLElement): Array<{ top: number; bottom: number }> {
  const range = block.ownerDocument.createRange();
  range.selectNodeContents(block);
  const rects = [...range.getClientRects()].filter((rect) => rect.width > 0 && rect.height > 0);
  rects.sort((a, b) => a.top - b.top);
  const lines: Array<{ top: number; bottom: number }> = [];
  for (const rect of rects) {
    const last = lines[lines.length - 1];
    // Misma línea si se superponen en más de la mitad del alto.
    if (last && rect.top < last.bottom - Math.min(rect.height, last.bottom - last.top) / 2) {
      last.bottom = Math.max(last.bottom, rect.bottom);
      last.top = Math.min(last.top, rect.top);
    } else {
      lines.push({ top: rect.top, bottom: rect.bottom });
    }
  }
  return lines;
}

/** El primer carácter que cae en el renglón que empieza en «top» (o null). */
function firstCharAt(block: HTMLElement, top: number): { node: Text; offset: number } | null {
  const doc = block.ownerDocument;
  const walker = doc.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  const range = doc.createRange();
  const charTop = (node: Text, i: number) => {
    range.setStart(node, i);
    range.setEnd(node, i + 1);
    const rects = [...range.getClientRects()].filter((rect) => rect.height > 0);
    return rects.length ? rects[rects.length - 1]!.top : -Infinity;
  };
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    const length = node.data.length;
    if (!length || charTop(node, length - 1) < top - 1) continue;
    let lo = 0;
    let hi = length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (charTop(node, mid) >= top - 1) hi = mid;
      else lo = mid + 1;
    }
    return { node, offset: lo };
  }
  return null;
}

/**
 * Parte un párrafo en el renglón que cruza el pie de la hoja, como Word: quedan al menos dos renglones de cada lado
 * (viudas y huérfanas). La primera parte no lleva espacio después y, si es justificado, su último renglón también se
 * justifica; la segunda no repite la sangría, el espacio antes ni el número de lista. Devuelve la segunda parte, o null
 * si no conviene partirlo (entonces pasa entero).
 */
function splitParagraph(block: HTMLElement, excess: number): HTMLElement | null {
  if (block.tagName !== "P") return null;
  const lines = linesOf(block);
  if (lines.length < 4) return null;
  const style = getComputedStyle(block);
  const bottom = block.getBoundingClientRect().bottom + (Number.parseFloat(style.marginBottom) || 0) - excess;
  let keep = lines.filter((line) => line.bottom <= bottom + 0.5).length;
  if (keep > lines.length - 2) keep = lines.length - 2;
  if (keep < 2) return null;
  const at = firstCharAt(block, lines[keep]!.top);
  if (!at) return null;
  const range = block.ownerDocument.createRange();
  range.setStart(at.node, at.offset);
  range.setEnd(block, block.childNodes.length);
  const tail = block.cloneNode(false) as HTMLElement;
  tail.removeAttribute("id");
  tail.appendChild(range.extractContents());
  tail.classList.add("pliego-sigue");
  tail.style.textIndent = "0";
  tail.style.marginTop = "0";
  block.style.marginBottom = "0";
  if (style.textAlign === "justify") block.style.textAlignLast = "justify";
  return tail;
}

/** Como Word: el primer párrafo de la página no lleva el espacio anterior (un título de capítulo arranca en el margen). */
function suppressTopSpace(page: HTMLElement) {
  const article = childrenOf(page, "article").find((item) => item.children.length > 0);
  const first = article?.firstElementChild as HTMLElement | null | undefined;
  if (!first || first.tagName !== "P" || first.style.marginTop === "0px") return;
  if ((Number.parseFloat(getComputedStyle(first).marginTop) || 0) > 0) first.style.marginTop = "0px";
}

/** Cada página lleva el encabezado y el pie que le tocan por sección y por página par o impar. */
function dress(root: HTMLElement, map: PageMap) {
  pagesOf(root).forEach((item, index) => {
    if (map.first.has(item)) return;
    const slots = map.templates.get(map.section.get(item) ?? 0);
    const model = slots?.[index % 2] ?? slots?.[(index + 1) % 2];
    if (!model) return;
    swap(item, "header", model.header);
    swap(item, "footer", model.footer);
  });
}

function pageNumbers(root: HTMLElement, set: Set<HTMLElement>): number[] {
  const pages = pagesOf(root);
  return [...set].map((item) => pages.indexOf(item) + 1).filter((n) => n > 0).sort((a, b) => a - b);
}

/**
 * Lo que no entra en la hoja pasa a una página nueva, igual que haría Word: se mueven párrafos o tablas enteros desde
 * el final, con sus notas al pie. Después cada página lleva el encabezado y el pie que le tocan por sección y por
 * página par o impar. Devuelve las páginas que siguen sin entrar (un solo bloque más alto que la hoja).
 */
export function repaginate(root: HTMLElement, map: PageMap): { added: number; tooTall: number[] } {
  const tooTall = new Set<HTMLElement>();
  const added = flow(root, map, tooTall);
  if (added) dress(root, map);
  // Numeración común de Word (seguida desde 1); numberFootnotes la rehace con la del documento.
  numberFootnotes(root, DEFAULT_RENDER_INFO, map);
  return { added, tooTall: pageNumbers(root, tooTall) };
}

/* ---------- Diseño de página que el visor no aplica: márgenes simétricos, encuadernación, secciones impar/par ---------- */

export type Padding = { left: number; right: number; top: number };

/**
 * Márgenes de una página (en píxeles) a partir de los del Word. Con márgenes simétricos el izquierdo del Word es el
 * interior: en las páginas pares (índice 1, 3…) pasa a la derecha. La encuadernación se suma al margen interior; sin
 * márgenes simétricos va siempre a la izquierda, o arriba si el Word lo pide.
 */
export function layoutPadding(
  base: Padding,
  index: number,
  gutterPt: number,
  mirrorMargins: boolean,
  gutterAtTop: boolean,
): Padding {
  const even = index % 2 === 1;
  const gutter = Math.max(0, gutterPt) * PT;
  let { left, right, top } = base;
  if (mirrorMargins && even) [left, right] = [right, left];
  if (gutter > 0) {
    if (mirrorMargins) {
      if (even) right += gutter;
      else left += gutter;
    } else if (gutterAtTop) top += gutter;
    else left += gutter;
  }
  return { left, right, top };
}

/** Los márgenes que puso el visor, guardados la primera vez (las páginas nuevas los heredan al copiarse). */
function basePadding(page: HTMLElement): Padding {
  const saved = page.dataset.pliegoPad?.split(" ").map(Number);
  if (saved && saved.length === 3 && saved.every(Number.isFinite)) {
    return { left: saved[0]!, right: saved[1]!, top: saved[2]! };
  }
  const style = getComputedStyle(page);
  const pad = {
    left: Number.parseFloat(style.paddingLeft) || 0,
    right: Number.parseFloat(style.paddingRight) || 0,
    top: Number.parseFloat(style.paddingTop) || 0,
  };
  page.dataset.pliegoPad = `${pad.left} ${pad.right} ${pad.top}`;
  return pad;
}

function setMargins(root: HTMLElement, map: PageMap, info: RenderInfo) {
  const pages = pagesOf(root);
  // Se leen todos antes de escribir, para no medir a mitad de cambio.
  const bases = pages.map(basePadding);
  pages.forEach((page, index) => {
    const layout = info.sections[map.section.get(page) ?? 0];
    const pad = layoutPadding(bases[index]!, index, layout?.gutterPt ?? 0, info.mirrorMargins, info.gutterAtTop);
    page.style.paddingLeft = `${pad.left}px`;
    page.style.paddingRight = `${pad.right}px`;
    page.style.paddingTop = `${pad.top}px`;
  });
}

export type PlanPage = { section: number; first: boolean };

/** 1 = tiene que empezar en página impar, 0 = en par, null = donde caiga. */
function wantedParity(layout: SectionLayout | undefined): 0 | 1 | null {
  if (layout?.type === "oddPage") return 1;
  if (layout?.type === "evenPage") return 0;
  return null;
}

/**
 * Antes de qué páginas (índices) va una página en blanco para que cada sección «página impar» o «página par» empiece
 * con un número de esa paridad. Los números se cuentan como numberPages; como Word, se mira el número de la página
 * (si la sección vuelve a empezar en 2 y pide impar, igual lleva la página en blanco).
 */
export function planBlanks(pages: PlanPage[], numbering: SectionNumbering[], layouts: SectionLayout[]): number[] {
  const out: number[] = [];
  let current = 0;
  let previous: number | null = null;
  pages.forEach((page, index) => {
    if (index === 0 || (page.first && page.section !== previous)) {
      const start = numbering[page.section]?.start;
      let n = start !== undefined ? start : current + 1;
      const parity = wantedParity(layouts[page.section]);
      if (index > 0 && parity !== null && n % 2 !== parity) {
        out.push(index);
        // La página en blanco se lleva el número siguiente de la sección anterior.
        if (start === undefined) n += 1;
      }
      current = n;
    } else {
      current += 1;
    }
    previous = page.section;
  });
  return out;
}

/** Secciones «impar»/«par» que en el PDF (página 1 = derecha) quedan del otro lado. */
export function wrongSide(pages: PlanPage[], layouts: SectionLayout[]): Array<{ section: number; page: number }> {
  const out: Array<{ section: number; page: number }> = [];
  let previous: number | null = null;
  pages.forEach((page, index) => {
    if (index > 0 && page.first && page.section !== previous) {
      const parity = wantedParity(layouts[page.section]);
      if (parity !== null && (index + 1) % 2 !== parity) out.push({ section: page.section, page: index + 1 });
    }
    previous = page.section;
  });
  return out;
}

function planOf(root: HTMLElement, map: PageMap, withBlanks: boolean): { pages: HTMLElement[]; plan: PlanPage[] } {
  const pages = pagesOf(root).filter((page) => withBlanks || !isBlank(page));
  return {
    pages,
    plan: pages.map((page) => ({ section: map.section.get(page) ?? 0, first: map.first.has(page) })),
  };
}

/** Una hoja vacía con la forma y el encabezado y pie de la página anterior (Word los imprime en la página en blanco). */
function blankAfter(previous: HTMLElement): HTMLElement {
  const blank = previous.cloneNode(false) as HTMLElement;
  blank.dataset.pliegoBlank = "";
  const header = childrenOf(previous, "header")[0];
  const article = childrenOf(previous, "article")[0];
  const footer = childrenOf(previous, "footer")[0];
  if (header) blank.appendChild(header.cloneNode(true));
  if (article) blank.appendChild(article.cloneNode(false));
  if (footer) blank.appendChild(footer.cloneNode(true));
  return blank;
}

/** Pone (o saca) las páginas en blanco que hacen falta. Devuelve si cambió algo. */
function placeBlanks(root: HTMLElement, map: PageMap, info: RenderInfo, numbering: SectionNumbering[]): boolean {
  const { pages, plan } = planOf(root, map, false);
  const wanted = new Set(planBlanks(plan, numbering, info.sections).map((index) => pages[index]!));
  const existing = pagesOf(root).filter(isBlank);
  const have = new Set(existing.map((blank) => blank.nextElementSibling as HTMLElement | null));
  if (wanted.size === have.size && [...wanted].every((page) => have.has(page))) return false;
  for (const blank of existing) {
    map.section.delete(blank);
    blank.remove();
  }
  for (const page of wanted) {
    const previous = page.previousElementSibling as HTMLElement | null;
    if (!previous?.matches(PAGE_SELECTOR)) continue;
    const blank = blankAfter(previous);
    previous.after(blank);
    map.section.set(blank, map.section.get(previous) ?? 0);
  }
  return true;
}

export type LayoutResult = {
  /** Páginas agregadas porque, con los márgenes del Word, el texto ya no entraba. */
  added: number;
  /** Páginas en blanco agregadas para que una sección empiece en página impar o par. */
  blanks: number;
  /** Páginas que siguen sin entrar (número de página en el PDF), mirando todas: reemplaza el tooTall de repaginate. */
  tooTall: number[];
  warnings: string[];
};

/**
 * Aplica lo que el visor no hace: márgenes simétricos, margen de encuadernación y secciones que empiezan en página
 * impar o par (con su página en blanco). Va después de repaginate y antes de numberPages. Al cambiar los márgenes el
 * texto se reacomoda, así que se vuelve a repaginar hasta que todo quede quieto.
 */
export function applyLayout(
  root: HTMLElement,
  map: PageMap,
  info: RenderInfo,
  numbering: SectionNumbering[] = [],
): LayoutResult {
  const warnings: string[] = [];
  const tooTall = new Set<HTMLElement>();
  const needsMargins = info.mirrorMargins || info.sections.some((layout) => (layout.gutterPt ?? 0) > 0);
  const needsParity = info.sections.some((layout) => wantedParity(layout) !== null);
  let added = 0;
  if (needsMargins || needsParity) {
    let settled = false;
    for (let round = 0; round < 6 && !settled; round += 1) {
      const changed = needsParity ? placeBlanks(root, map, info, numbering) : false;
      if (needsMargins) setMargins(root, map, info);
      const more = flow(root, map, tooTall);
      added += more;
      if (changed || more) dress(root, map);
      // Una vuelta sin páginas nuevas ni cambios en las en blanco: listo.
      settled = !changed && !more;
    }
    if (!settled) {
      warnings.push(
        "No pude dejar quietas las páginas en blanco de las secciones que empiezan en página impar o par: revisá en la vista previa que cada capítulo empiece del lado que corresponde.",
      );
    }
  }
  // Se vuelve a mirar todo al final: vale también cuando no hubo nada que cambiar.
  tooTall.clear();
  for (const page of pagesOf(root)) if (overflows(page, sheetHeight(page))) tooTall.add(page);
  if (needsParity) {
    const { plan } = planOf(root, map, true);
    for (const item of wrongSide(plan, info.sections)) {
      const side = wantedParity(info.sections[item.section]) === 1 ? "derecha (impar)" : "izquierda (par)";
      warnings.push(
        `La sección ${item.section + 1} tiene que empezar en página ${side}, pero en el PDF cae en la página ${item.page}, del otro lado: pasa cuando el Word vuelve a numerar esa sección desde un número que no coincide. Revisá la numeración de esa sección en el Word.`,
      );
    }
  }
  const blanks = pagesOf(root).filter(isBlank).length;
  return { added, blanks, tooTall: pageNumbers(root, tooTall), warnings };
}

/* ---------- Notas al pie ---------- */

export type NotePage = { section: number; custom: boolean[] };

/**
 * El número de cada llamada (null = nota con marca propia, que no lleva número ni cuenta). Seguida en todo el
 * documento, o de nuevo desde el inicio en cada sección o en cada página, como diga el Word.
 */
export function planFootnotes(
  pages: NotePage[],
  restart: RenderInfo["footnoteRestart"],
  startAt: number,
): Array<Array<number | null>> {
  const start = Number.isFinite(startAt) && startAt >= 1 ? Math.floor(startAt) : 1;
  let next = start;
  let previous: number | null = null;
  return pages.map((page) => {
    if (restart === "eachPage" || (restart === "eachSect" && previous !== null && page.section !== previous)) next = start;
    previous = page.section;
    return page.custom.map((custom) => (custom ? null : next++));
  });
}

/**
 * Numera las llamadas a nota y la lista de notas de cada página (ol start y el número de cada nota). Las notas con
 * marca propia quedan sin número: el símbolo ya está en el texto. Va al final, cuando ya no se mueven páginas.
 */
export function numberFootnotes(root: HTMLElement, info: RenderInfo, map: PageMap) {
  const pages = pagesOf(root);
  const refs = pages.map((page) => [...page.querySelectorAll(NOTE_REF)] as HTMLElement[]);
  const plan = planFootnotes(
    pages.map((page, index) => ({
      section: map.section.get(page) ?? 0,
      custom: refs[index]!.map((sup) => sup.dataset.pliegoCustom !== undefined),
    })),
    info.footnoteRestart,
    info.footnoteStart,
  );
  pages.forEach((page, index) => {
    const sups = refs[index]!;
    if (!sups.length) return;
    const numbers = plan[index]!;
    const notes = notesList(page);
    const items = notes ? ([...notes.children] as HTMLElement[]) : [];
    sups.forEach((sup, i) => {
      const n = numbers[i] ?? null;
      const item = items[i] as HTMLLIElement | undefined;
      sup.textContent = n === null ? "" : String(n);
      if (!item) return;
      if (n === null) {
        item.style.listStyleType = "none";
        item.removeAttribute("value");
      } else {
        item.style.listStyleType = "";
        item.value = n;
      }
    });
    const firstNumber = numbers.find((n): n is number => n !== null);
    if (notes instanceof HTMLOListElement && firstNumber !== undefined) notes.start = firstNumber;
  });
}

function roman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
    [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
  ];
  let out = "";
  for (const [value, digits] of table) while (n >= value) {
    out += digits;
    n -= value;
  }
  return out;
}

function letters(n: number): string {
  // Como Word: a…z, después aa…zz.
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

export function formatPageNumber(n: number, format?: string): string {
  switch (format) {
    case "lowerRoman":
      return roman(n);
    case "upperRoman":
      return roman(n).toUpperCase();
    case "lowerLetter":
      return letters(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    default:
      return String(n);
  }
}

/** Escribe en cada página su número real, con el inicio y el formato de su sección (pgNumType) o del campo (\* roman). */
export function numberPages(root: HTMLElement, sections: SectionNumbering[], map: PageMap) {
  const pages = pagesOf(root);
  let current = 0;
  let previous: number | null = null;
  const numbers = pages.map((page, index) => {
    const section = map.section.get(page) ?? 0;
    if (index === 0 || (map.first.has(page) && section !== previous)) {
      const start = sections[section]?.start;
      current = start !== undefined ? start : current + 1;
    } else {
      current += 1;
    }
    previous = section;
    return current;
  });
  const sectionTotals = new Map<number, number>();
  for (const page of pages) {
    const section = map.section.get(page) ?? 0;
    sectionTotals.set(section, (sectionTotals.get(section) ?? 0) + 1);
  }
  const pagePattern = new RegExp(`${MARK_OPEN}P(?::(\\w+))?${MARK_CLOSE}`, "g");
  for (const node of markTexts(root)) {
    const index = pages.findIndex((page) => page.contains(node));
    const section = index >= 0 ? (map.section.get(pages[index]!) ?? 0) : 0;
    node.nodeValue = (node.nodeValue ?? "")
      .replace(pagePattern, (_, format?: string) =>
        index >= 0 ? formatPageNumber(numbers[index]!, format ?? sections[section]?.format) : "",
      )
      .split(PAGES_MARK)
      .join(String(pages.length))
      .split(SECTION_PAGES_MARK)
      .join(String(sectionTotals.get(section) ?? pages.length));
  }
}

import { FOOTNOTE_MARK, MARK_CLOSE, MARK_OPEN, PAGES_MARK, SECTION_PAGES_MARK, type SectionNumbering } from "./fields.ts";
import {
  BREAK_MARK,
  DEFAULT_RENDER_INFO,
  HEADER_ROW_MARK,
  PROBE_MARK,
  suppressesSpaceAfterManualBreak,
  type RenderInfo,
  type SectionLayout,
} from "./render-copy.ts";

const PAGE_SELECTOR = "section.docx";
const NOTE_REF = "sup[data-pliego-note]";
/** Nota con marca propia (w:customMarkFollows): la llamada no lleva número, el símbolo ya está en el texto. */
export const CUSTOM_NOTE_MARK = `${MARK_OPEN}C${MARK_CLOSE}`;
/** Puntos a píxeles CSS. */
const PT = 96 / 72;
/** Guion opcional: donde se corta el renglón se dibuja; al final de un bloque ya no. */
const SHY = "­";

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

/** El último bloque con texto de la página. */
function lastBlockOf(page: HTMLElement): HTMLElement | null {
  const articles = childrenOf(page, "article").filter((article) => article.children.length > 0);
  return (articles[articles.length - 1]?.lastElementChild as HTMLElement | null) ?? null;
}

/**
 * Lo que Word no cuenta al pie de la página: el espacio después del último párrafo y el interlineado extra debajo de
 * su último renglón (Word lo pone arriba del texto; el navegador lo reparte arriba y abajo).
 */
function bottomSlack(page: HTMLElement): number {
  const last = lastBlockOf(page);
  if (!last) return 0;
  const style = getComputedStyle(last);
  let slack = Number.parseFloat(style.marginBottom) || 0;
  if (last.tagName === "P") {
    const line = Number.parseFloat(style.lineHeight);
    const size = Number.parseFloat(style.fontSize);
    if (Number.isFinite(line) && Number.isFinite(size)) slack += Math.max(0, (line - size * 1.17) / 2);
  }
  return slack;
}

function overflows(page: HTMLElement, height: number): boolean {
  return height > 0 && page.offsetHeight > height + 1 + bottomSlack(page);
}

/**
 * Como Word, el último párrafo de cada página no lleva espacio después (así la hoja no crece por ese margen). Va al
 * final, cuando las páginas ya no cambian.
 */
export function trimPageBottoms(root: HTMLElement) {
  for (const page of pagesOf(root)) {
    const last = lastBlockOf(page);
    if (last && (Number.parseFloat(getComputedStyle(last).marginBottom) || 0) > 0) last.style.marginBottom = "0px";
  }
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

/**
 * La raya que separa el texto de las notas al pie, como la dibuja Word: 2 pulgadas (144 pt) desde el borde de la
 * columna, a la altura del renglón del separador (cuyo alto ya reserva el estilo de la lista de notas). Va después de
 * paginar; no ocupa lugar. `height`: alto del renglón del separador, en px de la hoja.
 */
export function drawNoteSeparators(root: HTMLElement, height: number): number {
  if (!(height > 0)) return 0;
  let drawn = 0;
  for (const page of pagesOf(root)) {
    for (const old of childrenOf(page, "div").filter((div) => div.classList.contains("pliego-separador"))) old.remove();
    const notes = notesList(page);
    if (!notes || !notes.querySelector("li")) continue;
    const column = childrenOf(page, "article")[0] ?? notes;
    if (getComputedStyle(page).position === "static") page.style.position = "relative";
    const line = page.ownerDocument.createElement("div");
    line.className = "pliego-separador";
    line.style.cssText =
      `position:absolute;left:${column.offsetLeft}px;top:${(notes.offsetTop + height * 0.6).toFixed(2)}px;` +
      "width:144pt;height:0;border-top:0.5pt solid #000;margin:0;padding:0;pointer-events:none;";
    page.insertBefore(line, notes);
    drawn += 1;
  }
  return drawn;
}

/** Página en blanco que se agregó para que una sección empiece en página impar o par. */
function isBlank(page: HTMLElement): boolean {
  return page.dataset.pliegoBlank !== undefined;
}

function blockCount(page: HTMLElement): number {
  return childrenOf(page, "article").reduce((sum, article) => sum + article.children.length, 0);
}

function lastBlock(page: HTMLElement): HTMLElement | null {
  const articles = childrenOf(page, "article").filter((article) => article.children.length > 0);
  return (articles[articles.length - 1]?.lastElementChild as HTMLElement | null) ?? null;
}

/**
 * Con varias partes de texto en la página (secciones continuas, notas finales), solo la última empuja hacia abajo las
 * notas al pie y el pie de página; las otras quedan pegadas a la siguiente. No cambia cuánto entra en la página.
 */
function settleArticles(root: HTMLElement) {
  for (const page of pagesOf(root)) {
    const articles = childrenOf(page, "article");
    articles.forEach((article, index) => {
      const wanted = index === articles.length - 1 ? "" : "0px";
      if (article.style.marginBottom !== wanted) article.style.marginBottom = wanted;
    });
  }
}

export type PageMap = {
  /** Sección de cada página. */
  section: Map<HTMLElement, number>;
  /** Páginas donde empieza una sección (el visor eligió para ellas el encabezado de primera página). */
  first: Set<HTMLElement>;
  /** Encabezado y pie del Word para cada sección, según la página sea impar (0) o par (1). */
  templates: Map<number, Array<{ header: HTMLElement | null; footer: HTMLElement | null } | undefined>>;
  /** Las notas finales, sacadas de la página de muestra: repaginate las pone al final del texto. */
  endnotes: HTMLElement | null;
};

/** La llamada a nota que sigue a una marca: se saltean los textos que son solo marcas (o quedaron vacíos). */
function supAfter(node: Text): HTMLElement | null {
  const onlyMarks = new RegExp(`^(?:${MARK_OPEN}[^${MARK_CLOSE}]*${MARK_CLOSE})*$`);
  let next = node.nextSibling;
  while (next && next.nodeType === Node.TEXT_NODE && onlyMarks.test(next.nodeValue ?? "")) next = next.nextSibling;
  return next instanceof HTMLElement && next.tagName === "SUP" ? next : null;
}

/**
 * Quita del texto las marcas de sección, de notas al pie, de saltos manuales y de filas de título, y anota de qué
 * sección es cada página y qué encabezado y pie le tocan (para repetirlos bien en las páginas que se agreguen). Las
 * páginas de muestra (render-copy.ts) solo sirven para eso: se sacan, y las notas finales que el visor puso en la última
 * quedan aparte hasta repaginate.
 */
export function takeMarks(root: HTMLElement): PageMap {
  const pages = pagesOf(root);
  const section = new Map<HTMLElement, number>();
  const first = new Set<HTMLElement>();
  const starts = new Map<HTMLElement, number>();
  const probes = new Set<HTMLElement>();
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
    const page = node.parentElement?.closest<HTMLElement>(PAGE_SELECTOR) ?? null;
    for (const match of text.matchAll(pattern)) {
      if (page && !starts.has(page)) starts.set(page, Number(match[1]));
    }
    if (page && text.includes(PROBE_MARK)) probes.add(page);
    if (text.includes(BREAK_MARK)) {
      const paragraph = node.parentElement?.closest("p");
      if (paragraph) paragraph.dataset.pliegoCont = "";
    }
    if (text.includes(HEADER_ROW_MARK)) {
      const row = node.parentElement?.closest("tr");
      if (row) row.dataset.pliegoHeader = "";
    }
    for (const mark of [FOOTNOTE_MARK, CUSTOM_NOTE_MARK, BREAK_MARK, HEADER_ROW_MARK, PROBE_MARK]) {
      text = text.split(mark).join("");
    }
    node.nodeValue = text.replace(pattern, "");
  }
  // Una marca de nota propia sin su marca de nota común igual es una llamada (por si llegan en textos separados).
  for (const sup of root.querySelectorAll("sup[data-pliego-custom]")) (sup as HTMLElement).dataset.pliegoNote = "";
  const templates: PageMap["templates"] = new Map();
  let endnotes: HTMLElement | null = null;
  let current = 0;
  pages.forEach((page, index) => {
    const begins = starts.get(page);
    if (begins !== undefined) current = begins;
    if (probes.has(page)) {
      // En las muestras no hay notas al pie: una lista ahí son las notas finales (el visor las pone en la última).
      for (const list of childrenOf(page, "ol")) endnotes = list;
    } else {
      section.set(page, current);
      // Cómo empieza la página: con la sección, o después de un salto de página manual (las de flow, «natural»).
      page.dataset.pliegoBreak = begins !== undefined ? "section" : "manual";
      if (begins !== undefined) first.add(page);
      for (const article of childrenOf(page, "article")) article.dataset.pliegoSection = String(current);
    }
    if (begins !== undefined) return;
    // Las páginas que no empiezan sección (las muestras incluidas) traen el encabezado y el pie común o de página par
    // que eligió el visor según la paridad con que las dibujó.
    const slots = templates.get(current) ?? [];
    slots[index % 2] ??= { header: childrenOf(page, "header")[0] ?? null, footer: childrenOf(page, "footer")[0] ?? null };
    templates.set(current, slots);
  });
  for (const page of probes) page.remove();
  hangBreakLines(pagesOf(root));
  return { section, first, templates, endnotes };
}

/**
 * Un párrafo que empieza con un salto de página manual (un título «ÍNDICE» con el salto adelante) queda partido en dos:
 * lo de antes del salto, sin nada visible, al final de la página anterior, y el texto en la página nueva. En Word ese
 * renglón del salto cuelga al pie de la página sin ocupar lugar (ni su espacio anterior): si la página ya estaba llena,
 * no se agrega una página en blanco. Acá ese pedazo vacío se saca.
 */
function hangBreakLines(pages: HTMLElement[]) {
  for (let index = 1; index < pages.length; index += 1) {
    const page = pages[index]!;
    if (page.dataset.pliegoBreak !== "manual") continue;
    const lead = childrenOf(page, "article").find((article) => article.children.length > 0)?.firstElementChild as HTMLElement | null;
    if (!lead || lead.tagName !== "P" || lead.dataset.pliegoCont === undefined) continue;
    const articles = childrenOf(pages[index - 1]!, "article").filter((article) => article.children.length > 0);
    const before = articles[articles.length - 1]?.lastElementChild as HTMLElement | null | undefined;
    if (!before || before.tagName !== "P" || before.className !== lead.className) continue;
    if ((before.textContent ?? "").trim() || before.querySelector("img, svg, object, canvas")) continue;
    // Fuera de la página (si quedara escondido, la paginación lo seguiría tomando por el último bloque); sus
    // marcadores (los destinos del índice) pasan al principio del párrafo de la página nueva.
    const anchors = [...before.querySelectorAll<HTMLElement>("[id]")];
    lead.prepend(...anchors);
    before.remove();
  }
}

/**
 * Con «páginas por salto» el visor empieza una página en cada sección, también en las continuas. Como en Word, una
 * sección continua sigue en la misma página que la anterior (si la hoja es del mismo tamaño): su texto y sus notas
 * pasan a la página anterior. Va después de takeMarks y antes de repaginate. Devuelve cuántas páginas juntó.
 */
export function mergeContinuous(root: HTMLElement, map: PageMap, info: RenderInfo): number {
  let merged = 0;
  for (const page of pagesOf(root)) {
    if (!map.first.has(page)) continue;
    if (info.sections[map.section.get(page) ?? 0]?.type !== "continuous") continue;
    const previous = page.previousElementSibling as HTMLElement | null;
    if (!previous?.matches(PAGE_SELECTOR) || isBlank(previous)) continue;
    if (previous.style.width !== page.style.width || previous.style.minHeight !== page.style.minHeight) continue;
    const before = notesList(previous) ?? childrenOf(previous, "footer")[0] ?? null;
    for (const article of childrenOf(page, "article")) previous.insertBefore(article, before);
    const notes = notesList(page);
    if (notes) {
      const own = notesList(previous);
      if (own) own.append(...[...notes.children]);
      else previous.insertBefore(notes, childrenOf(previous, "footer")[0] ?? null);
    }
    map.section.delete(page);
    map.first.delete(page);
    page.remove();
    merged += 1;
  }
  if (merged) settleArticles(root);
  return merged;
}

/**
 * Las notas finales van al final del texto de la última página, como en Word, una lista por nota (con su número) para
 * que puedan seguir en las páginas siguientes de a una.
 */
function placeEndnotes(root: HTMLElement, map: PageMap) {
  const list = map.endnotes as HTMLOListElement | null;
  map.endnotes = null;
  if (!list) return;
  const pages = pagesOf(root).filter((page) => !isBlank(page));
  const page = pages[pages.length - 1];
  if (!page) return;
  const article = root.ownerDocument.createElement("article");
  article.className = "pliego-notas-finales";
  article.dataset.pliegoSection = String(map.section.get(page) ?? 0);
  const items = [...list.children] as HTMLElement[];
  const start = Number.isFinite(list.start) ? list.start : 1;
  items.forEach((item, index) => {
    const one = list.cloneNode(false) as HTMLOListElement;
    one.start = start + index;
    // Entre una nota y la siguiente no hay más espacio que el de la propia nota.
    if (index > 0) one.style.marginTop = "0px";
    if (index < items.length - 1) one.style.marginBottom = "0px";
    one.appendChild(item);
    article.appendChild(one);
  });
  const articles = childrenOf(page, "article");
  const last = articles[articles.length - 1];
  if (last) last.after(article);
  else page.insertBefore(article, notesList(page) ?? childrenOf(page, "footer")[0] ?? null);
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

/* ---------- Pasar lo que no entra a la página siguiente ---------- */

/** La página nueva que recibe lo que no entra, con los lugares donde va cada cosa. */
type Fresh = { page: HTMLElement; articles: Map<HTMLElement, HTMLElement>; notes: HTMLElement | null };

function freshAfter(current: HTMLElement, map: PageMap): Fresh {
  const page = current.cloneNode(false) as HTMLElement;
  delete page.dataset.pliegoBlank;
  // Empieza por un salto natural: como en Word, su primer párrafo no lleva el espacio anterior.
  page.dataset.pliegoBreak = "natural";
  const header = childrenOf(current, "header")[0];
  const footer = childrenOf(current, "footer")[0];
  if (header) page.appendChild(header.cloneNode(true));
  if (footer) page.appendChild(footer.cloneNode(true));
  current.after(page);
  map.section.set(page, map.section.get(current) ?? 0);
  return { page, articles: new Map(), notes: null };
}

/** La copia, en la página nueva, de la parte de texto de donde sale un bloque. */
function targetFor(fresh: Fresh, source: HTMLElement): HTMLElement {
  let target = fresh.articles.get(source);
  if (!target) {
    target = source.cloneNode(false) as HTMLElement;
    fresh.articles.set(source, target);
    const page = fresh.page;
    const before = childrenOf(page, "article")[0] ?? childrenOf(page, "ol")[0] ?? childrenOf(page, "footer")[0];
    page.insertBefore(target, before ?? null);
  }
  return target;
}

/** Las notas de las llamadas que se movieron son las últimas de la lista de esta página. */
function moveNotes(current: HTMLElement, fresh: Fresh, count: number) {
  const notes = notesList(current);
  if (!count || !notes) return;
  if (!fresh.notes) {
    fresh.notes = notes.cloneNode(false) as HTMLElement;
    fresh.page.insertBefore(fresh.notes, childrenOf(fresh.page, "footer")[0] ?? null);
  }
  fresh.notes.prepend(...[...notes.children].slice(-count));
  if (!notes.children.length) notes.remove();
}

/** Saca las partes de texto que quedaron vacías (siempre queda al menos una). */
function dropEmptyArticles(page: HTMLElement) {
  const articles = childrenOf(page, "article");
  const empty = articles.filter((article) => !article.children.length);
  if (empty.length === articles.length) empty.shift();
  for (const article of empty) article.remove();
}

/** Hasta dónde puede llegar el texto de una página que no entra (coordenadas de pantalla), y la escala de la vista. */
function contentLimit(page: HTMLElement, height: number): { limit: number; k: number } | null {
  const articles = childrenOf(page, "article").filter((article) => article.children.length > 0);
  const last = articles[articles.length - 1];
  if (!last || !(page.offsetHeight > 0)) return null;
  const box = page.getBoundingClientRect();
  if (!(box.height > 0)) return null;
  const k = box.height / page.offsetHeight;
  // Lo que va debajo del texto (notas, pie, margen) ocupa lo mismo cuando la página no entra en la hoja.
  const below = box.bottom - last.getBoundingClientRect().bottom;
  return { limit: box.top + height * k - below, k };
}

/**
 * El último bloque que puede quedar en la página (índice, o -1 si ni el primero entra): su borde de abajo no pasa el
 * límite, contando el lugar que liberan las notas de los bloques que se van. Los bordes crecen de arriba abajo (texto
 * en una columna); refs da las llamadas a nota de cada bloque y relief(n) el lugar que dejan las últimas n notas.
 */
export function lastFitting(
  count: number,
  bottom: (index: number) => number,
  refs: (index: number) => number,
  totalRefs: number,
  limit: number,
  relief: (notes: number) => number,
): number {
  // Lo que está más abajo que el límite con todas las notas fuera no puede quedar: se busca a saltos.
  const widest = limit + relief(totalRefs) + 0.5;
  let lo = 0;
  let hi = count;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (bottom(mid) > widest) hi = mid;
    else lo = mid + 1;
  }
  let after = totalRefs;
  for (let index = 0; index < lo; index += 1) after -= refs(index);
  for (let index = lo - 1; index >= 0; index -= 1) {
    if (bottom(index) <= limit + relief(Math.max(0, after)) + 0.5) return index;
    after += refs(index);
  }
  return -1;
}

/**
 * Los bloques que seguro no entran en la página, medidos de una sola vez: todos los que siguen al primero que no
 * entra (ese queda, porque a lo mejor se parte). Con texto en columnas no se calcula (los bloques no van uno debajo
 * del otro) y se pasan de a uno.
 */
function surelyOut(page: HTMLElement, height: number): HTMLElement[] {
  const articles = childrenOf(page, "article").filter((article) => article.children.length > 0);
  if (articles.some((article) => (Number.parseInt(getComputedStyle(article).columnCount, 10) || 1) > 1)) return [];
  const blocks = articles.flatMap((article) => [...article.children] as HTMLElement[]);
  if (blocks.length < 3) return [];
  const area = contentLimit(page, height);
  if (!area) return [];
  const { limit, k } = area;
  const notes = notesList(page);
  const items = notes ? ([...notes.children] as HTMLElement[]) : [];
  const relief = (n: number): number => {
    if (!notes || n <= 0 || !items.length) return 0;
    if (n >= items.length) {
      const style = getComputedStyle(notes);
      const margins = (Number.parseFloat(style.marginTop) || 0) + (Number.parseFloat(style.marginBottom) || 0);
      return notes.getBoundingClientRect().height + margins * k;
    }
    return notes.getBoundingClientRect().bottom - items[items.length - n]!.getBoundingClientRect().top;
  };
  const counts = new Map<number, number>();
  const refs = (index: number): number => {
    let found = counts.get(index);
    if (found === undefined) {
      found = blocks[index]!.querySelectorAll(NOTE_REF).length;
      counts.set(index, found);
    }
    return found;
  };
  // Cada llamada de la página tiene su nota en la lista: no hace falta recorrer todo el texto para contarlas.
  const totalRefs = items.length;
  const bottom = (index: number): number => {
    const block = blocks[index]!;
    // Sin el espacio después: si este bloque queda último, Word no lo cuenta.
    return block.getBoundingClientRect().bottom;
  };
  const keep = lastFitting(blocks.length, bottom, refs, totalRefs, limit, relief);
  return blocks.slice(keep + 2);
}

/** Las filas de título del principio de la tabla (se repiten arriba en cada página). */
function headerRows(table: HTMLTableElement): HTMLTableRowElement[] {
  const out: HTMLTableRowElement[] = [];
  for (const row of table.rows) {
    if (row.dataset.pliegoHeader === undefined) break;
    out.push(row);
  }
  return out;
}

/** Cuántas filas tienen que quedar arriba como mínimo: las de título y una más (si hay de título). */
function minRows(table: HTMLTableElement): number {
  const titles = headerRows(table).length;
  return titles > 0 && titles < table.rows.length ? titles + 1 : 1;
}

function canSplit(page: HTMLElement): boolean {
  const block = lastBlock(page);
  if (!block) return false;
  if (block.tagName === "P") return linesOf(block).length >= 4;
  if (block.tagName === "TABLE") return (block as HTMLTableElement).rows.length > minRows(block as HTMLTableElement);
  return false;
}

export type SpanCell = { row: number; col: number; span: number };

/**
 * Celdas combinadas en vertical que cruzan el corte de una tabla (las primeras «headRows» filas quedan arriba): cuántas
 * filas de la combinación quedan arriba y cuántas siguen en la página nueva.
 */
export function spansAcross<T extends SpanCell>(cells: T[], headRows: number): Array<T & { keep: number; rest: number }> {
  return cells
    .filter((cell) => cell.span > 1 && cell.row < headRows && cell.row + cell.span > headRows)
    .map((cell) => ({ ...cell, keep: headRows - cell.row, rest: cell.row + cell.span - headRows }));
}

/** Las celdas de una fila con su columna (contando las celdas ocultas de las combinaciones, como el visor). */
function cellColumns(row: HTMLTableRowElement): Array<{ cell: HTMLTableCellElement; col: number }> {
  const out: Array<{ cell: HTMLTableCellElement; col: number }> = [];
  let col = 0;
  for (const cell of row.cells) {
    out.push({ cell, col });
    col += Math.max(1, cell.colSpan);
  }
  return out;
}

/**
 * El visor dibuja una celda combinada en vertical como una sola celda alta, y en las filas que siguen deja celdas
 * ocultas. Si la tabla se corta a mitad de la combinación, arriba la celda se acorta y abajo la oculta de la primera
 * fila pasa a ser la celda que sigue (si no, las celdas de esas filas se correrían una columna).
 */
function splitMerges(head: HTMLTableElement, tail: HTMLTableElement, offset: number) {
  const first = tail.rows[offset];
  if (!first) return;
  const cells: Array<SpanCell & { cell: HTMLTableCellElement }> = [];
  [...head.rows].forEach((row, index) => {
    for (const { cell, col } of cellColumns(row)) {
      if (cell.style.display !== "none") cells.push({ row: index, col, span: cell.rowSpan, cell });
    }
  });
  const twins = cellColumns(first);
  for (const item of spansAcross(cells, head.rows.length)) {
    item.cell.rowSpan = item.keep;
    const twin = twins.find((entry) => entry.col === item.col)?.cell;
    if (twin && twin.style.display === "none") {
      twin.style.display = "";
      twin.rowSpan = item.rest;
    }
  }
}

/**
 * Una tabla que no entra se parte por filas: las que sobran van a una copia de la tabla en la página nueva, con las
 * filas de título repetidas arriba (como Word). Primero pasan de una vez las que seguro no entran.
 */
function splitTable(table: HTMLTableElement, page: HTMLElement, height: number, target: HTMLElement): HTMLTableElement {
  const tail = table.cloneNode(false) as HTMLTableElement;
  tail.removeAttribute("id");
  for (const part of [...table.children]) if (part.tagName === "COLGROUP") tail.appendChild(part.cloneNode(true));
  const keep = minRows(table);
  const titles = keep > 1 ? headerRows(table) : [];
  for (const row of titles) {
    const copy = row.cloneNode(true) as HTMLTableRowElement;
    // La repetición no lleva llamadas a nota: la nota ya va con la fila de arriba.
    for (const sup of copy.querySelectorAll(NOTE_REF)) sup.remove();
    copy.dataset.pliegoRepeat = "";
    tail.appendChild(copy);
  }
  target.insertBefore(tail, target.firstChild);
  let firstMoved: HTMLTableRowElement | null = null;
  const area = contentLimit(page, height);
  if (area) {
    const rows = [...table.rows];
    let lo = Math.max(0, keep - 1);
    let hi = rows.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (rows[mid]!.getBoundingClientRect().bottom > area.limit + 0.5) hi = mid;
      else lo = mid + 1;
    }
    const out = rows.slice(lo + 1);
    for (const row of out) tail.appendChild(row);
    firstMoved = out[0] ?? null;
  }
  while (overflows(page, height) && table.rows.length > keep) {
    const row = table.rows[table.rows.length - 1]!;
    if (firstMoved) firstMoved.before(row);
    else tail.appendChild(row);
    firstMoved = row;
  }
  splitMerges(table, tail, titles.length);
  return tail;
}

/**
 * Pasa a páginas nuevas lo que no entra en la hoja, igual que haría Word: se mueven párrafos o tablas enteros desde
 * el final, con sus notas al pie, o se parten. Se puede llamar otra vez (después de cambiar márgenes): solo toca las
 * páginas que no entran. Devuelve cuántas páginas agregó y anota las que siguen sin entrar (un solo bloque más alto que
 * la hoja que no se puede partir).
 */
function flow(root: HTMLElement, map: PageMap, tooTall: Set<HTMLElement>, info: RenderInfo): number {
  let added = 0;
  let page = root.querySelector(PAGE_SELECTOR) as HTMLElement | null;
  while (page) {
    const current = page;
    suppressTopSpace(current, info);
    const height = sheetHeight(current);
    if (overflows(current, height) && (blockCount(current) > 1 || canSplit(current))) {
      const fresh = freshAfter(current, map);
      added += 1;
      // Lo que seguro no entra pasa de una vez (una sola medición); después se ajusta bloque por bloque.
      const bulk = surelyOut(current, height);
      if (bulk.length) {
        let refs = 0;
        for (let index = bulk.length - 1; index >= 0; index -= 1) {
          const block = bulk[index]!;
          refs += block.querySelectorAll(NOTE_REF).length;
          const target = targetFor(fresh, block.parentElement as HTMLElement);
          target.insertBefore(block, target.firstChild);
        }
        moveNotes(current, fresh, refs);
        dropEmptyArticles(current);
      }
      while (overflows(current, height) && (blockCount(current) > 1 || canSplit(current))) {
        const last = lastBlock(current)!;
        const target = targetFor(fresh, last.parentElement as HTMLElement);
        let block: HTMLElement;
        if (last.tagName === "TABLE" && (last as HTMLTableElement).rows.length > minRows(last as HTMLTableElement)) {
          block = splitTable(last as HTMLTableElement, current, height, target);
        } else {
          const tail = splitParagraph(last, current.offsetHeight - height);
          // El único bloque de la página, que no se puede partir: queda (y la página se avisa como demasiado alta).
          if (!tail && blockCount(current) === 1) break;
          block = tail ?? last;
          target.insertBefore(block, target.firstChild);
        }
        moveNotes(current, fresh, block.querySelectorAll(NOTE_REF).length);
        dropEmptyArticles(current);
      }
      if (blockCount(fresh.page) === 0) {
        map.section.delete(fresh.page);
        fresh.page.remove();
        added -= 1;
      } else {
        dropEmptyArticles(fresh.page);
        // La página nueva es de la sección de su primer texto (después de una sección continua puede ser otra).
        const lead = childrenOf(fresh.page, "article").find((article) => article.children.length > 0);
        const tagged = Number.parseInt(lead?.dataset.pliegoSection ?? "", 10);
        if (Number.isFinite(tagged)) map.section.set(fresh.page, tagged);
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

/** El carácter justo antes del corte (puede estar en el texto anterior), para ver si es un guion opcional. */
function charBefore(block: HTMLElement, at: { node: Text; offset: number }): { node: Text; offset: number } | null {
  if (at.offset > 0) return { node: at.node, offset: at.offset - 1 };
  const walker = block.ownerDocument.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let previous: Text | null = null;
  for (let node = walker.nextNode() as Text | null; node && node !== at.node; node = walker.nextNode() as Text | null) {
    if (node.data.length) previous = node;
  }
  return previous ? { node: previous, offset: previous.data.length - 1 } : null;
}

/**
 * Dónde empieza la segunda parte cuando el corte cae al principio de un texto: el elemento más alto del que ese texto
 * es lo primero (una llamada a nota pasa entera, sin dejar arriba una copia vacía que se numeraría).
 */
function cutStart(block: HTMLElement, at: { node: Text; offset: number }): Node | null {
  if (at.offset !== 0) return null;
  const leads = (node: Node) => {
    for (let sibling = node.previousSibling; sibling; sibling = sibling.previousSibling) {
      if (!(sibling.nodeType === Node.TEXT_NODE && !(sibling as Text).data)) return false;
    }
    return true;
  };
  let node: Node = at.node;
  while (node.parentNode && node.parentNode !== block && leads(node)) node = node.parentNode;
  return node;
}

/**
 * Parte un párrafo en el renglón que cruza el pie de la hoja, como Word: quedan al menos dos renglones de cada lado
 * (viudas y huérfanas). La primera parte no lleva espacio después y, si es justificado, su último renglón también se
 * justifica; la segunda no repite la sangría, el espacio antes ni el número de lista. Si el renglón de arriba termina
 * en un guion opcional, el guion queda escrito (al final del bloque ya no se dibujaría). Devuelve la segunda parte, o
 * null si no conviene partirlo (entonces pasa entero).
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
  const before = charBefore(block, at);
  if (before && before.node.data[before.offset] === SHY) before.node.replaceData(before.offset, 1, "-");
  const range = block.ownerDocument.createRange();
  const start = cutStart(block, at);
  if (start) range.setStartBefore(start);
  else range.setStart(at.node, at.offset);
  range.setEnd(block, block.childNodes.length);
  const tail = block.cloneNode(false) as HTMLElement;
  tail.removeAttribute("id");
  delete tail.dataset.pliegoCont;
  tail.appendChild(range.extractContents());
  tail.classList.add("pliego-sigue");
  tail.style.textIndent = "0";
  tail.style.marginTop = "0";
  block.style.marginBottom = "0";
  if (style.textAlign === "justify") block.style.textAlignLast = "justify";
  return tail;
}

/** Si el primer párrafo de una página lleva el espacio anterior, según cómo empieza la página. */
export function keepsSpaceBefore(start: string | undefined, continuation: boolean, info: RenderInfo): boolean {
  // La continuación de un párrafo que cortó un salto manual no es el comienzo del párrafo.
  if (continuation) return false;
  // Salto natural (la página la agregó la paginación).
  if (start === "natural") return false;
  if (start === "manual") return !suppressesSpaceAfterManualBreak(info);
  // Comienzo de sección (o no se sabe): queda, como en Word.
  return true;
}

/**
 * Como Word: el primer párrafo de la página no lleva el espacio anterior después de un salto natural, ni (desde Word
 * 2013) después de un salto manual. Al comienzo de una sección nueva queda (la caída de un título de capítulo).
 */
function suppressTopSpace(page: HTMLElement, info: RenderInfo) {
  const article = childrenOf(page, "article").find((item) => item.children.length > 0);
  const first = article?.firstElementChild as HTMLElement | null | undefined;
  if (!first || first.tagName !== "P" || first.style.marginTop === "0px") return;
  if (keepsSpaceBefore(page.dataset.pliegoBreak, first.dataset.pliegoCont !== undefined, info)) return;
  if ((Number.parseFloat(getComputedStyle(first).marginTop) || 0) > 0) first.style.marginTop = "0px";
}

/**
 * Cada página lleva el encabezado y el pie que le tocan por sección y por página par o impar. La primera página de
 * una sección con «primera página distinta» queda con el de primera página que puso el visor.
 */
function dress(root: HTMLElement, map: PageMap, info: RenderInfo) {
  pagesOf(root).forEach((item, index) => {
    const section = map.section.get(item) ?? 0;
    if (map.first.has(item)) {
      const layout = info.sections[section];
      // Sin datos de la sección no se sabe si es primera página distinta: queda la que eligió el visor.
      if (!layout || layout.titlePage) return;
    }
    const slots = map.templates.get(section);
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
 * el final, con sus notas al pie, o se parten. Las notas finales van al final del texto. Después cada página lleva el
 * encabezado y el pie que le tocan por sección y por página par o impar. Va después de takeMarks y mergeContinuous.
 * Devuelve las páginas que siguen sin entrar (un solo bloque más alto que la hoja).
 */
export function repaginate(
  root: HTMLElement,
  map: PageMap,
  info: RenderInfo = DEFAULT_RENDER_INFO,
): { added: number; tooTall: number[] } {
  const tooTall = new Set<HTMLElement>();
  placeEndnotes(root, map);
  const added = flow(root, map, tooTall, info);
  dress(root, map, info);
  numberFootnotes(root, info, map);
  settleArticles(root);
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

/**
 * La encuadernación ya viene sumada al margen en la copia para el visor (render-copy.ts): acá solo se invierten
 * izquierda y derecha en las páginas pares con márgenes simétricos, que no cambia el ancho del texto.
 */
function setMargins(root: HTMLElement, info: RenderInfo) {
  const pages = pagesOf(root);
  // Se leen todos antes de escribir, para no medir a mitad de cambio.
  const bases = pages.map(basePadding);
  pages.forEach((page, index) => {
    const pad = layoutPadding(bases[index]!, index, 0, info.mirrorMargins, info.gutterAtTop);
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
 * Aplica lo que el visor no hace: márgenes simétricos y secciones que empiezan en página impar o par (con su página
 * en blanco). La encuadernación ya viene en los márgenes de la copia. Va después de repaginate y antes de numberPages.
 * Al cambiar las páginas en blanco cambia qué página es par, así que se vuelve a mirar hasta que todo quede quieto.
 */
export function applyLayout(
  root: HTMLElement,
  map: PageMap,
  info: RenderInfo,
  numbering: SectionNumbering[] = [],
): LayoutResult {
  const warnings: string[] = [];
  const tooTall = new Set<HTMLElement>();
  const needsMargins = info.mirrorMargins;
  const needsParity = info.sections.some((layout) => wantedParity(layout) !== null);
  let added = 0;
  if (needsMargins || needsParity) {
    let settled = false;
    for (let round = 0; round < 6 && !settled; round += 1) {
      const changed = needsParity ? placeBlanks(root, map, info, numbering) : false;
      if (needsMargins) setMargins(root, info);
      const more = flow(root, map, tooTall, info);
      added += more;
      if (changed || more) dress(root, map, info);
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
  settleArticles(root);
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
      // El número dentro del texto de la nota (donde Word tiene w:footnoteRef), con el estilo de llamada de nota.
      const slot = item.querySelector<HTMLElement>("[data-pliego-note-number]");
      if (slot) slot.textContent = n === null ? "" : String(n);
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

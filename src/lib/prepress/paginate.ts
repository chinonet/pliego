import { FOOTNOTE_MARK, MARK_CLOSE, MARK_OPEN, PAGES_MARK, SECTION_PAGES_MARK, type SectionNumbering } from "./fields.ts";

const PAGE_SELECTOR = "section.docx";
const NOTE_REF = "sup[data-pliego-note]";

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

export type PageMap = {
  /** Sección de cada página. */
  section: Map<HTMLElement, number>;
  /** Páginas donde empieza una sección (el visor eligió para ellas el encabezado de primera página). */
  first: Set<HTMLElement>;
  /** Encabezado y pie del Word para cada sección, según la página sea impar (0) o par (1). */
  templates: Map<number, Array<{ header: HTMLElement | null; footer: HTMLElement | null } | undefined>>;
};

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
  for (const node of markTexts(root)) {
    let text = node.nodeValue ?? "";
    for (const match of text.matchAll(pattern)) {
      const page = pages.find((item) => item.contains(node));
      if (page && !starts.has(page)) starts.set(page, Number(match[1]));
    }
    text = text.replace(pattern, "");
    if (text.includes(FOOTNOTE_MARK)) {
      const sup = node.nextSibling;
      if (sup instanceof HTMLElement && sup.tagName === "SUP") sup.dataset.pliegoNote = "";
      text = text.split(FOOTNOTE_MARK).join("");
    }
    node.nodeValue = text;
  }
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
 * Lo que no entra en la hoja pasa a una página nueva, igual que haría Word: se mueven párrafos o tablas enteros desde
 * el final, con sus notas al pie. Después cada página lleva el encabezado y el pie que le tocan por sección y por
 * página par o impar. Devuelve las páginas que siguen sin entrar (un solo bloque más alto que la hoja).
 */
export function repaginate(root: HTMLElement, map: PageMap): { added: number; tooTall: number[] } {
  let added = 0;
  const tooTall = new Set<HTMLElement>();
  let page = root.querySelector(PAGE_SELECTOR) as HTMLElement | null;
  while (page) {
    const current = page;
    const height = sheetHeight(current);
    const blocks = () => childrenOf(current, "article").reduce((sum, article) => sum + article.children.length, 0);
    if (overflows(current, height) && blocks() > 1) {
      const fresh = current.cloneNode(false) as HTMLElement;
      const header = childrenOf(current, "header")[0];
      const footer = childrenOf(current, "footer")[0];
      if (header) fresh.appendChild(header.cloneNode(true));
      if (footer) fresh.appendChild(footer.cloneNode(true));
      current.after(fresh);
      map.section.set(fresh, map.section.get(current) ?? 0);
      added += 1;
      const freshArticles = new Map<HTMLElement, HTMLElement>();
      let freshNotes: HTMLElement | null = null;
      while (overflows(current, height) && blocks() > 1) {
        const articles = childrenOf(current, "article").filter((article) => article.children.length > 0);
        const source = articles[articles.length - 1]!;
        let target = freshArticles.get(source);
        if (!target) {
          target = source.cloneNode(false) as HTMLElement;
          freshArticles.set(source, target);
          const before = childrenOf(fresh, "article")[0] ?? childrenOf(fresh, "ol")[0] ?? childrenOf(fresh, "footer")[0];
          fresh.insertBefore(target, before ?? null);
        }
        const block = source.lastElementChild as HTMLElement;
        target.insertBefore(block, target.firstChild);
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
    const following = current.nextElementSibling as HTMLElement | null;
    page = following?.matches(PAGE_SELECTOR) ? following : null;
  }

  const pages = pagesOf(root);
  if (added) {
    pages.forEach((item, index) => {
      if (map.first.has(item)) return;
      const slots = map.templates.get(map.section.get(item) ?? 0);
      const model = slots?.[index % 2] ?? slots?.[(index + 1) % 2];
      if (!model) return;
      swap(item, "header", model.header);
      swap(item, "footer", model.footer);
    });
  }
  // Las llamadas a nota se numeran por página, como las numera el visor.
  for (const item of pages) {
    [...item.querySelectorAll(NOTE_REF)].forEach((sup, index) => {
      sup.textContent = String(index + 1);
    });
  }
  return { added, tooTall: [...tooTall].map((item) => pages.indexOf(item) + 1).filter((n) => n > 0) };
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

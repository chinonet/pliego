import { MARK_CLOSE, MARK_OPEN, PAGE_MARK, PAGES_MARK, type SectionNumbering } from "./fields.ts";

const PAGE_SELECTOR = "section.docx";

function pagesOf(root: HTMLElement): HTMLElement[] {
  return [...root.querySelectorAll(PAGE_SELECTOR)] as HTMLElement[];
}

/** Alto de la hoja según el Word (el visor lo pone como alto mínimo; si el contenido no entra, la página crece). */
function sheetHeight(page: HTMLElement): number {
  const value = Number.parseFloat(getComputedStyle(page).minHeight);
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function overflows(page: HTMLElement, height: number): boolean {
  return height > 0 && page.offsetHeight > height + 1;
}

function pick(page: HTMLElement | null, tag: string): HTMLElement | null {
  if (!page) return null;
  return ([...page.children] as HTMLElement[]).find((child) => child.tagName.toLowerCase() === tag) ?? null;
}

/**
 * Lo que no entra en la hoja pasa a una página nueva, igual que haría Word: se mueven párrafos o tablas enteros desde
 * el final. La página nueva lleva el encabezado y el pie de la sección. Devuelve las páginas que siguen sin entrar
 * (un solo bloque más alto que la hoja).
 */
export function repaginate(root: HTMLElement): { added: number; tooTall: number[] } {
  let added = 0;
  const tooTall = new Set<HTMLElement>();
  let page = root.querySelector(PAGE_SELECTOR) as HTMLElement | null;
  while (page) {
    const height = sheetHeight(page);
    if (overflows(page, height)) {
      const articles = ([...page.children] as HTMLElement[]).filter((child) => child.tagName === "ARTICLE");
      const source = articles[articles.length - 1];
      if (source && source.children.length > 1) {
        const next = page.nextElementSibling as HTMLElement | null;
        const model = next?.matches(PAGE_SELECTOR) ? next : page;
        const fresh = page.cloneNode(false) as HTMLElement;
        const header = pick(model, "header") ?? pick(page, "header");
        const footer = pick(model, "footer") ?? pick(page, "footer");
        const article = source.cloneNode(false) as HTMLElement;
        if (header) fresh.appendChild(header.cloneNode(true));
        fresh.appendChild(article);
        if (footer) fresh.appendChild(footer.cloneNode(true));
        page.after(fresh);
        added += 1;
        while (overflows(page, height) && source.children.length > 1) {
          article.insertBefore(source.lastElementChild!, article.firstChild);
        }
      } else {
        tooTall.add(page);
      }
    }
    const following = page.nextElementSibling as HTMLElement | null;
    page = following?.matches(PAGE_SELECTOR) ? following : null;
  }
  const pages = pagesOf(root);
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

function markTexts(root: HTMLElement): Text[] {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const found: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeValue?.includes(MARK_OPEN)) found.push(node as Text);
  }
  return found;
}

/** Quita las marcas de sección del texto. Devuelve, por página, la primera sección que empieza en ella. */
export function takeSectionMarks(root: HTMLElement): Map<HTMLElement, number> {
  const pages = pagesOf(root);
  const starts = new Map<HTMLElement, number>();
  const pattern = new RegExp(`${MARK_OPEN}S(\\d+)${MARK_CLOSE}`, "g");
  for (const node of markTexts(root)) {
    const text = node.nodeValue ?? "";
    for (const match of text.matchAll(pattern)) {
      const page = pages.find((item) => item.contains(node));
      if (page && !starts.has(page)) starts.set(page, Number(match[1]));
    }
    node.nodeValue = text.replace(pattern, "");
  }
  return starts;
}

/** Escribe en cada página su número real, con el inicio y el formato de su sección (pgNumType). */
export function numberPages(root: HTMLElement, sections: SectionNumbering[], starts: Map<HTMLElement, number>) {
  const pages = pagesOf(root);
  let section = 0;
  let current = 0;
  const numbers = pages.map((page, index) => {
    const begins = starts.get(page);
    if (index === 0 || begins !== undefined) {
      section = begins ?? section;
      const start = sections[section]?.start;
      current = start !== undefined ? start : current + 1;
    } else {
      current += 1;
    }
    return formatPageNumber(current, sections[section]?.format);
  });
  const total = String(pages.length);
  for (const node of markTexts(root)) {
    const index = pages.findIndex((page) => page.contains(node));
    node.nodeValue = (node.nodeValue ?? "")
      .split(PAGE_MARK)
      .join(index >= 0 ? numbers[index] : "")
      .split(PAGES_MARK)
      .join(total);
  }
}

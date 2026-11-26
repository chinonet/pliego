import { MARK_CLOSE } from "./fields.ts";
import { EMPTY_PARAGRAPH_MARK, NOTE_NUMBER_MARK, NOTE_SEPARATOR_MARK, TABS_MARK_OPEN } from "./render-copy.ts";

/**
 * Lo que render-copy deja escrito para después de dibujar: la letra de la marca de párrafo de cada párrafo vacío y
 * las tabulaciones de los párrafos con tabulaciones a la derecha, al centro o decimales.
 */

export type TabStop = { pos: number; align: string; leader: string };

/** «359.55:right:dot;…» → tabulaciones (posición en puntos desde el borde de la columna). */
export function parseTabStops(payload: string | undefined): TabStop[] {
  if (!payload) return [];
  const stops: TabStop[] = [];
  for (const part of payload.split(";")) {
    const [pos, align, leader] = part.split(":");
    const value = Number.parseFloat(pos ?? "");
    if (!Number.isFinite(value) || !align) continue;
    stops.push({ pos: value, align, leader: leader || "none" });
  }
  return stops.sort((a, b) => a.pos - b.pos);
}

/**
 * Lee y saca las marcas apenas se dibuja el Word, antes de que el visor calcule las tabulaciones (medio segundo
 * después): con el texto de la marca adentro, las calcularía corridas. Va después de arreglar las reglas de estilo, para
 * leer la letra que de verdad tiene la marca de párrafo.
 */
export function takeRenderMarks(root: HTMLElement): { empty: number; tabs: number } {
  const doc = root.ownerDocument;
  const view = doc.defaultView;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const found: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? "";
    if ([TABS_MARK_OPEN, EMPTY_PARAGRAPH_MARK, NOTE_SEPARATOR_MARK, NOTE_NUMBER_MARK].some((mark) => text.includes(mark))) {
      found.push(node as Text);
    }
  }
  let empty = 0;
  let tabs = 0;
  for (const node of found) {
    const holder = node.parentElement;
    const paragraph = holder?.closest("p") ?? null;
    let text = node.nodeValue ?? "";
    if (text.includes(NOTE_NUMBER_MARK)) {
      // El lugar del número de la nota en su propio texto: lo escribe numberFootnotes (paginate.ts).
      if (holder) holder.dataset.pliegoNoteNumber = "";
      node.nodeValue = text.split(NOTE_NUMBER_MARK).join("");
      continue;
    }
    if (text.includes(NOTE_SEPARATOR_MARK)) {
      // La muestra del separador de notas: queda con un espacio (un renglón de su letra) hasta medirla.
      if (paragraph) paragraph.dataset.pliegoNoteSeparator = "";
      node.nodeValue = text.split(NOTE_SEPARATOR_MARK).join(" ");
      continue;
    }
    if (text.includes(EMPTY_PARAGRAPH_MARK)) {
      if (paragraph && holder && view) {
        const style = view.getComputedStyle(holder);
        paragraph.dataset.pliegoMarkFont = JSON.stringify({ list: style.fontFamily, size: Number.parseFloat(style.fontSize) });
        empty += 1;
      }
      text = text.split(EMPTY_PARAGRAPH_MARK).join("");
    }
    for (let start = text.indexOf(TABS_MARK_OPEN); start >= 0; start = text.indexOf(TABS_MARK_OPEN)) {
      const end = text.indexOf(MARK_CLOSE, start);
      if (end < 0) break;
      if (paragraph) {
        paragraph.dataset.pliegoTabs = text.slice(start + TABS_MARK_OPEN.length, end);
        tabs += 1;
      }
      text = text.slice(0, start) + text.slice(end + MARK_CLOSE.length);
    }
    node.nodeValue = text;
    if (text) continue;
    // El tramo de la marca queda vacío: se saca (y los que lo envolvían, si quedan vacíos), sin tocar el párrafo.
    let element: HTMLElement | null = holder;
    node.remove();
    while (element && element !== paragraph && !element.childNodes.length) {
      const up: HTMLElement | null = element.parentElement;
      element.remove();
      element = up;
    }
  }
  return { empty, tabs };
}

/**
 * Mide la muestra del separador de notas (ya con el interlineado de Word) y la saca. Devuelve su alto en px de la
 * hoja (0 si el libro no tiene notas al pie). Va después de fixLineSpacing y antes de paginar.
 */
export function takeNoteSeparator(root: HTMLElement): number {
  const probe = root.querySelector<HTMLElement>("p[data-pliego-note-separator]");
  if (!probe) return 0;
  const view = root.ownerDocument.defaultView;
  const style = view?.getComputedStyle(probe);
  const height =
    probe.offsetHeight + (Number.parseFloat(style?.marginTop ?? "") || 0) + (Number.parseFloat(style?.marginBottom ?? "") || 0);
  probe.remove();
  return height > 0 ? height : 0;
}

/** La letra de la marca de párrafo que guardó takeRenderMarks, o null. */
export function markFontOf(paragraph: HTMLElement): { list: string; size: number } | null {
  const raw = paragraph.dataset.pliegoMarkFont;
  if (!raw) return null;
  try {
    const value = JSON.parse(raw) as { list?: unknown; size?: unknown };
    if (typeof value.list === "string" && typeof value.size === "number" && value.size > 0) {
      return { list: value.list, size: value.size };
    }
  } catch {
    /* Marca dañada: se ignora. */
  }
  return null;
}

/** Ancho (en px de pantalla) del texto que va después de una tabulación, hasta la siguiente o el final del párrafo. */
function followingWidth(tab: HTMLElement, paragraph: HTMLElement, next: HTMLElement | undefined): number {
  const doc = paragraph.ownerDocument;
  const walker = doc.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  const range = doc.createRange();
  let width = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (tab.contains(node)) continue;
    if (!(tab.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)) continue;
    if (next && (next.contains(node) || next.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)) break;
    range.selectNodeContents(node);
    for (const rect of range.getClientRects()) width += rect.width;
  }
  return width;
}

function setLeader(tab: HTMLElement, leader: string) {
  tab.style.textDecoration = "inherit";
  tab.style.removeProperty("text-decoration-style");
  if (leader === "dot" || leader === "middleDot") {
    tab.style.textDecoration = "underline";
    tab.style.textDecorationStyle = "dotted";
  } else if (leader === "hyphen" || leader === "heavy" || leader === "underscore") {
    tab.style.textDecoration = "underline";
  }
}

/**
 * Rehace las tabulaciones a la derecha, al centro y decimales como Word. El visor mide lo que sigue a la tabulación con
 * un rectángulo que, si ese texto ya pasó al renglón de abajo, abarca los dos renglones, y da un ancho equivocado (en un
 * índice, el número de página quedaba partido). Word: el texto termina en la tabulación (o se centra en ella); si no
 * entra entre lo anterior y la tabulación, va pegado a lo anterior, y si tampoco entra en el renglón, la tabulación y lo
 * que sigue pasan al renglón siguiente (con su relleno de puntos desde el principio). Las tabulaciones a la izquierda
 * quedan como las calculó el visor. Va después de que el visor calcula las suyas y antes de medir los renglones.
 */
export function fixRightTabs(root: HTMLElement): number {
  const view = root.ownerDocument.defaultView;
  if (!view) return 0;
  let changed = 0;
  for (const paragraph of root.querySelectorAll<HTMLElement>("p[data-pliego-tabs]")) {
    const stops = parseTabStops(paragraph.dataset.pliegoTabs).filter((stop) => stop.align !== "bar");
    if (!stops.length) continue;
    const tabs = [...paragraph.querySelectorAll<HTMLElement>(".docx-tab-stop")].filter((tab) => tab.closest("p") === paragraph);
    const style = view.getComputedStyle(paragraph);
    // Un párrafo partido entre dos columnas (CSS) tiene una caja que abarca las dos: las medidas no sirven. Queda como
    // lo calculó el visor.
    if (paragraph.getClientRects().length > 1) continue;
    tabs.forEach((tab, index) => {
      const before = tab.style.wordSpacing;
      const beforeText = tab.textContent;
      // Para medir, la tabulación es un punto de corte sin ancho: el visor pone un espacio que no corta, pegado a la
      // palabra anterior, y una palabra larga al final del renglón (el título de un índice) bajaba entera.
      tab.textContent = "​";
      tab.style.wordSpacing = "0px";
      const box = paragraph.getBoundingClientRect();
      const scale = box.width > 0 ? paragraph.offsetWidth / box.width : 1;
      const toPt = (screen: number) => screen * scale * 0.75;
      // Las posiciones de Word se miden desde el borde de la columna, no desde la sangría del párrafo.
      const columnLeft = box.left - (Number.parseFloat(style.marginLeft) || 0) / scale;
      const lineEnd = toPt(box.right - (Number.parseFloat(style.paddingRight) || 0) / scale - columnLeft);
      const lineStart = toPt(box.left + (Number.parseFloat(style.paddingLeft) || 0) / scale - columnLeft);
      const measure = () => {
        const rect = tab.getBoundingClientRect();
        return { left: toPt(rect.left - columnLeft), space: toPt(rect.width) };
      };
      let { left } = measure();
      let stop = stops.find((item) => item.pos > left + 0.01);
      if (!stop || stop.align === "left" || stop.align === "start" || stop.align === "num") {
        tab.textContent = beforeText;
        tab.style.wordSpacing = before;
        return;
      }
      const following = toPt(followingWidth(tab, paragraph, tabs[index + 1]));
      const share = (align: string) => (align === "center" ? 0.5 : align === "left" || align === "start" ? 0 : 1);
      let fill = stop.pos - share(stop.align) * following - left;
      if (fill < 0) {
        if (left + following <= lineEnd + 0.01) {
          // Entra en el renglón aunque no antes de la tabulación: va pegado a lo anterior, sin relleno.
          setLeader(tab, "none");
          changed += 1;
          return;
        }
        // Tampoco entra en el renglón: la tabulación y lo que sigue pasan al renglón de abajo. Solo si bajar sirve: si
        // la tabulación ya está al principio del renglón, o lo que sigue no entra ni en un renglón entero, queda pegado.
        const top = tab.getBoundingClientRect().top;
        if (left <= lineStart + 0.5 || following > lineEnd - lineStart + 0.01) {
          setLeader(tab, "none");
          changed += 1;
          return;
        }
        const br = paragraph.ownerDocument.createElement("br");
        br.dataset.pliegoTabBreak = "";
        tab.parentNode?.insertBefore(br, tab);
        ({ left } = measure());
        if (tab.getBoundingClientRect().top <= top + 0.5) {
          // No bajó (por ejemplo, ya estaba al principio de un renglón partido): sin salto.
          br.remove();
          setLeader(tab, "none");
          changed += 1;
          return;
        }
        stop = stops.find((item) => item.pos > left + 0.01) ?? stop;
        fill = stop.pos - share(stop.align) * following - left;
      }
      // El relleno: un espacio que no corta, estirado hasta la tabulación (un poco menos, para que el redondeo no
      // empuje lo que sigue al renglón de abajo).
      tab.textContent = " ";
      const space = measure().space;
      if (fill < space) {
        // No entra ni el espacio: queda sin ancho (lo que sigue termina a menos de un espacio de la tabulación).
        tab.textContent = "​";
        setLeader(tab, "none");
        changed += 1;
        return;
      }
      tab.style.wordSpacing = `${(fill - space - 0.05).toFixed(2)}pt`;
      setLeader(tab, stop.leader);
      changed += 1;
    });
  }
  return changed;
}

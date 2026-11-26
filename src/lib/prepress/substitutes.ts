import JSZip from "jszip";
import { fontKey } from "./font-face.ts";

type FaceStyle = "normal" | "italic";

/** Familias métricas de uso libre. Solo se aplican si el usuario lo pide. */
export const SUBSTITUTE_FACES: Array<{
  family: string;
  file: string;
  weight: number;
  style: FaceStyle;
}> = [
  { family: "Calibri", file: "/fonts/Carlito-Regular.ttf", weight: 400, style: "normal" },
  { family: "Calibri", file: "/fonts/Carlito-Bold.ttf", weight: 700, style: "normal" },
  { family: "Calibri", file: "/fonts/Carlito-Italic.ttf", weight: 400, style: "italic" },
  { family: "Calibri", file: "/fonts/Carlito-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Calibri Light", file: "/fonts/Carlito-Regular.ttf", weight: 300, style: "normal" },
  { family: "Cambria", file: "/fonts/Caladea-Regular.ttf", weight: 400, style: "normal" },
  { family: "Cambria", file: "/fonts/Caladea-Bold.ttf", weight: 700, style: "normal" },
  { family: "Cambria", file: "/fonts/Caladea-Italic.ttf", weight: 400, style: "italic" },
  { family: "Cambria", file: "/fonts/Caladea-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Arial", file: "/fonts/LiberationSans-Regular.ttf", weight: 400, style: "normal" },
  { family: "Arial", file: "/fonts/LiberationSans-Bold.ttf", weight: 700, style: "normal" },
  { family: "Arial", file: "/fonts/LiberationSans-Italic.ttf", weight: 400, style: "italic" },
  { family: "Arial", file: "/fonts/LiberationSans-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Helvetica", file: "/fonts/LiberationSans-Regular.ttf", weight: 400, style: "normal" },
  { family: "Helvetica", file: "/fonts/LiberationSans-Bold.ttf", weight: 700, style: "normal" },
  { family: "Helvetica", file: "/fonts/LiberationSans-Italic.ttf", weight: 400, style: "italic" },
  { family: "Helvetica", file: "/fonts/LiberationSans-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Regular.ttf", weight: 400, style: "normal" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Bold.ttf", weight: 700, style: "normal" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Italic.ttf", weight: 400, style: "italic" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Times", file: "/fonts/LiberationSerif-Regular.ttf", weight: 400, style: "normal" },
  { family: "Times", file: "/fonts/LiberationSerif-Bold.ttf", weight: 700, style: "normal" },
  { family: "Times", file: "/fonts/LiberationSerif-Italic.ttf", weight: 400, style: "italic" },
  { family: "Times", file: "/fonts/LiberationSerif-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Courier New", file: "/fonts/LiberationMono-Regular.ttf", weight: 400, style: "normal" },
  { family: "Courier New", file: "/fonts/LiberationMono-Bold.ttf", weight: 700, style: "normal" },
  { family: "Courier New", file: "/fonts/LiberationMono-Italic.ttf", weight: 400, style: "italic" },
  { family: "Courier New", file: "/fonts/LiberationMono-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Courier", file: "/fonts/LiberationMono-Regular.ttf", weight: 400, style: "normal" },
  { family: "Courier", file: "/fonts/LiberationMono-Bold.ttf", weight: 700, style: "normal" },
  { family: "Courier", file: "/fonts/LiberationMono-Italic.ttf", weight: 400, style: "italic" },
  { family: "Courier", file: "/fonts/LiberationMono-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Tahoma", file: "/fonts/Tahoma.ttf", weight: 400, style: "normal" },
  { family: "Tahoma", file: "/fonts/Tahoma-Bold.ttf", weight: 700, style: "normal" },
  { family: "Candara", file: "/fonts/Candara-Regular.ttf", weight: 400, style: "normal" },
  { family: "Candara", file: "/fonts/Candara-Bold.ttf", weight: 700, style: "normal" },
  { family: "Candara", file: "/fonts/Candara-Italic.ttf", weight: 400, style: "italic" },
  { family: "Candara", file: "/fonts/Candara-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Candara Light", file: "/fonts/Candara-Light.ttf", weight: 400, style: "normal" },
  { family: "Candara Light", file: "/fonts/Candara-LightItalic.ttf", weight: 400, style: "italic" },
  { family: "MS Gothic", file: "/fonts/MSGothic.ttf", weight: 400, style: "normal" },
  { family: "Symbol", file: "/fonts/Symbol.ttf", weight: 400, style: "normal" },
  { family: "Wingdings", file: "/fonts/Wingdings.ttf", weight: 400, style: "normal" },
];

/** Clave de una cara: familia (sin mayúsculas ni espacios) + negrita sí/no + cursiva sí/no. */
export function faceKey(family: string, weight: number, style: string): string {
  return `${fontKey(family)}|${weight >= 600 ? "b" : "r"}|${style === "normal" ? "n" : "i"}`;
}

const STYLE_ID = "pliego-substitutes";

/**
 * Nombre aparte con el que se publica también la sustituta de una cara que el Word trae recortada (solo las letras
 * que usa): puesto en segundo lugar en la lista de fuentes, completa lo que a esa cara le falta.
 */
export function substituteAlias(family: string): string {
  return `${family} (sustituta)`;
}

type SubstituteFace = (typeof SUBSTITUTE_FACES)[number];

/** Caras sustitutas que se publican; `skip` son claves de faceKey (por ejemplo, las que el Word trae enteras). */
function activeFaces(skip?: Iterable<string>) {
  const skipped = new Set(skip ?? []);
  return SUBSTITUTE_FACES.filter((face) => !skipped.has(faceKey(face.family, face.weight, face.style)));
}

/** Las sustitutas que, además, van con su nombre aparte (`fallback`: claves faceKey de caras recortadas del Word). */
function aliasFaces(fallback?: Iterable<string>): SubstituteFace[] {
  const wanted = new Set(fallback ?? []);
  return wanted.size ? SUBSTITUTE_FACES.filter((face) => wanted.has(faceKey(face.family, face.weight, face.style))) : [];
}

function faceRule(family: string, face: SubstituteFace): string {
  return `@font-face{font-family:"${family}";src:url("${face.file}") format("truetype");font-weight:${face.weight};font-style:${face.style};font-display:swap;}`;
}

/**
 * Publica las sustitutas. `skip`: caras que no van (las que el Word trae enteras). `fallback`: caras que el Word trae
 * recortadas; su sustituta va también con el nombre de substituteAlias, para completar las letras que falten.
 */
export function installSubstitutes(enabled: boolean, skip?: Iterable<string>, fallback?: Iterable<string>) {
  const existing = document.getElementById(STYLE_ID);
  if (!enabled) {
    existing?.remove();
    return;
  }
  const style = existing ?? document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = [
    ...activeFaces(skip).map((face) => faceRule(face.family, face)),
    ...aliasFaces(fallback).map((face) => faceRule(substituteAlias(face.family), face)),
  ].join("\n");
  document.head.appendChild(style);
}

export async function loadSubstituteFaces(skip?: Iterable<string>, fallback?: Iterable<string>) {
  if (!("fonts" in document)) return;
  const faces = [
    ...activeFaces(skip).map((face) => ({ family: face.family, face })),
    ...aliasFaces(fallback).map((face) => ({ family: substituteAlias(face.family), face })),
  ];
  await Promise.all(
    faces.map(({ family, face }) =>
      document.fonts.load(`${face.style} ${face.weight} 16px "${family}"`).catch(() => undefined),
    ),
  );
  await document.fonts.ready;
}

/**
 * La lista de fuentes con la sustituta en segundo lugar, detrás de la primera («"Calibri", serif» →
 * «"Calibri", "Calibri (sustituta)", serif»). Si ya está, queda igual.
 */
export function withSubstituteAlias(fontFamily: string, alias: string): string {
  const parts = fontFamily.split(",").map((part) => part.trim()).filter(Boolean);
  const bare = (part: string) => part.replace(/^['"]+|['"]+$/g, "").toLowerCase();
  if (parts.some((part) => bare(part) === alias.toLowerCase())) return fontFamily;
  const quoted = `"${alias.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
  return [...parts.slice(0, 1), quoted, ...parts.slice(1)].join(", ");
}

/**
 * Pone la sustituta detrás de cada texto escrito con una cara recortada del Word (`fallback`, claves faceKey), para
 * que la pantalla y el PDF saquen de ahí las letras que esa cara no trae (números de página, «‑», versalitas…) en vez
 * de una fuente del sistema. Solo toca el estilo del visor, nunca el Word. Devuelve con qué deshacerlo.
 */
export function addSubstituteFallback(root: ParentNode, fallback: Iterable<string>): () => void {
  const wanted = new Map<string, string>();
  for (const face of aliasFaces(fallback)) {
    wanted.set(faceKey(face.family, face.weight, face.style), substituteAlias(face.family));
  }
  const undo: Array<() => void> = [];
  if (!wanted.size) return () => undefined;
  const seen = new Set<Element>();
  const walker = document.createTreeWalker(root as Node, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = node.parentElement;
    if (!parent || seen.has(parent) || !(node.nodeValue ?? "").trim() || parent.closest("style, script")) continue;
    seen.add(parent);
    const css = getComputedStyle(parent);
    const family = firstFamily(css.fontFamily);
    const alias = wanted.get(faceKey(family, Number.parseInt(css.fontWeight, 10) || 400, css.fontStyle === "normal" ? "normal" : "italic"));
    if (!alias) continue;
    const next = withSubstituteAlias(css.fontFamily, alias);
    if (next === css.fontFamily) continue;
    const element = parent as HTMLElement;
    const before = element.style.getPropertyValue("font-family");
    const priority = element.style.getPropertyPriority("font-family");
    element.style.setProperty("font-family", next, priority);
    undo.push(() => {
      if (before) element.style.setProperty("font-family", before, priority);
      else element.style.removeProperty("font-family");
    });
  }
  return () => {
    for (const step of undo.reverse()) step();
  };
}

/* ---------- Caras que faltan ---------- */

export type UsedFace = { family: string; weight: number; style: FaceStyle; chars: number; sample: string };
export type AvailableFace = { family: string; weight: number | [number, number]; style: FaceStyle };
export type MissingFace = UsedFace & {
  /** «Calibri negrita cursiva». */
  label: string;
  /** No hay ninguna cara de esa familia (ni sustituta, ni subida, ni del Word). */
  familyMissing: boolean;
};

const GENERIC = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "math",
  "emoji",
  "fangsong",
  "ui-serif",
  "ui-sans-serif",
  "ui-monospace",
  "ui-rounded",
  "inherit",
  "initial",
]);

export function faceLabel(family: string, weight: number, style: string): string {
  const parts = [family];
  if (weight >= 600) parts.push("negrita");
  else if (weight < 350) parts.push("fina");
  if (style !== "normal") parts.push("cursiva");
  if (parts.length === 1) parts.push("normal");
  return parts.join(" ");
}

function firstFamily(list: string): string {
  return (list.split(",")[0] ?? "").trim().replace(/^['"]+|['"]+$/g, "");
}

function parseWeight(value: string): number | [number, number] {
  const words = value
    .trim()
    .split(/\s+/)
    .map((word) => (word === "bold" ? 700 : word === "normal" ? 400 : Number(word) || 400));
  return words.length > 1 ? [words[0]!, words[1]!] : words[0]!;
}

/**
 * Cruza lo que usa el texto con las caras reales que hay. Una cara sirve si coincide la cursiva y si es negrita
 * (600 o más) o no: así eligen el navegador y el PDF; sin la cara, el PDF no inventa la negrita ni la cursiva.
 */
export function missingFacesFrom(used: UsedFace[], available: AvailableFace[]): MissingFace[] {
  const families = new Set(available.map((face) => fontKey(face.family)));
  const keys = new Set<string>();
  for (const face of available) {
    const [min, max] = Array.isArray(face.weight) ? face.weight : [face.weight, face.weight];
    if (min < 600) keys.add(faceKey(face.family, 400, face.style));
    if (max >= 600) keys.add(faceKey(face.family, 700, face.style));
  }
  const missing = new Map<string, MissingFace>();
  for (const face of used) {
    const key = faceKey(face.family, face.weight, face.style);
    if (keys.has(key)) continue;
    const previous = missing.get(key);
    if (previous) {
      previous.chars += face.chars;
      if (previous.sample.length < 40) previous.sample = `${previous.sample} ${face.sample}`.trim().slice(0, 40);
      continue;
    }
    const weight = face.weight >= 600 ? 700 : 400;
    missing.set(key, {
      ...face,
      weight,
      label: faceLabel(face.family, weight, face.style),
      familyMissing: !families.has(fontKey(face.family)),
    });
  }
  return [...missing.values()].sort((a, b) => b.chars - a.chars);
}

/** Familia, peso y estilo de cada texto visible de las páginas (lo que calcula el navegador, no lo que dice el Word). */
export function usedTextFaces(root: Node): UsedFace[] {
  const seen = new Map<Element, UsedFace | null>();
  const faces = new Map<string, UsedFace>();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = (node.nodeValue ?? "").replace(/\s+/g, "");
    const parent = node.parentElement;
    if (!text || !parent || parent.closest("style, script")) continue;
    let face = seen.get(parent);
    if (face === undefined) {
      face = null;
      if (parent.getClientRects().length > 0) {
        const css = getComputedStyle(parent);
        const family = firstFamily(css.fontFamily);
        if (family && !GENERIC.has(family.toLowerCase())) {
          face = {
            family,
            weight: Number.parseInt(css.fontWeight, 10) || 400,
            style: css.fontStyle === "normal" ? "normal" : "italic",
            chars: 0,
            sample: "",
          };
        }
      }
      seen.set(parent, face);
    }
    if (!face) continue;
    const key = `${faceKey(face.family, face.weight, face.style)}|${face.family}`;
    const total = faces.get(key) ?? { ...face };
    total.chars += text.length;
    if (total.sample.length < 40) total.sample = `${total.sample} ${(node.nodeValue ?? "").trim()}`.trim().slice(0, 40);
    faces.set(key, total);
  }
  return [...faces.values()];
}

/** Caras declaradas con @font-face (sustitutas, subidas o incrustadas en el Word) que no fallaron al cargar. */
export function availableFaces(): AvailableFace[] {
  if (!("fonts" in document)) return [];
  return [...document.fonts]
    .filter((face) => face.status !== "error")
    .map((face) => ({
      family: firstFamily(face.family),
      weight: parseWeight(face.weight),
      style: face.style === "normal" ? ("normal" as const) : ("italic" as const),
    }));
}

/** Combinaciones familia + negrita + cursiva que usa el texto y no tienen una cara real: para pedirlas en la pantalla. */
export function missingFaces(root: Node): MissingFace[] {
  return missingFacesFrom(usedTextFaces(root), availableFaces());
}

/* ---------- Fuentes incrustadas en el Word ---------- */

const EMBED_TYPES: Record<string, { weight: number; style: FaceStyle }> = {
  embedRegular: { weight: 400, style: "normal" },
  embedBold: { weight: 700, style: "normal" },
  embedItalic: { weight: 400, style: "italic" },
  embedBoldItalic: { weight: 700, style: "italic" },
};

export type EmbeddedRef = {
  family: string;
  weight: number;
  style: FaceStyle;
  /** Word guardó solo las letras que usa el documento. */
  subsetted: boolean;
  /** Ruta dentro del .docx (word/fonts/font1.odttf). */
  path: string;
  fontKey: string;
};

export type EmbeddedFace = EmbeddedRef & {
  /** El archivo de la fuente ya descifrado (TTF/OTF), listo para @font-face o registerFont. */
  bytes: Uint8Array;
  /** La licencia de la fuente no permite incrustarla (OS/2 fsType). */
  restricted: boolean;
};

function xmlAttr(tag: string, name: string): string | undefined {
  return new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];
}

async function refsFromZip(zip: JSZip): Promise<EmbeddedRef[]> {
  const table = await zip.file("word/fontTable.xml")?.async("string");
  if (!table) return [];
  const relsXml = (await zip.file("word/_rels/fontTable.xml.rels")?.async("string")) ?? "";
  const targets = new Map<string, string>();
  for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = xmlAttr(tag, "Id");
    const target = xmlAttr(tag, "Target");
    if (!id || !target) continue;
    targets.set(id, target.startsWith("/") ? target.slice(1) : `word/${target.replace(/^\.\//, "")}`);
  }
  const refs: EmbeddedRef[] = [];
  for (const chunk of table.split(/<w:font\b/).slice(1)) {
    const family = xmlAttr(chunk.slice(0, chunk.indexOf(">") + 1), "w:name");
    if (!family) continue;
    for (const match of chunk.matchAll(/<w:(embedRegular|embedBold|embedItalic|embedBoldItalic)\b([^>]*)>/g)) {
      const kind = EMBED_TYPES[match[1]!]!;
      const tag = match[2] ?? "";
      const id = xmlAttr(tag, "r:id");
      const path = id ? targets.get(id) : undefined;
      if (!path) continue;
      refs.push({
        family,
        weight: kind.weight,
        style: kind.style,
        subsetted: /^(?:1|true|on)$/.test(xmlAttr(tag, "w:subsetted") ?? ""),
        path,
        fontKey: xmlAttr(tag, "w:fontKey") ?? "",
      });
    }
  }
  return refs;
}

/** Caras que el Word trae incrustadas (word/fontTable.xml, embedRegular/embedBold/… con su r:id). */
export async function embeddedFamilies(bytes: Uint8Array): Promise<EmbeddedRef[]> {
  try {
    return await refsFromZip(await JSZip.loadAsync(bytes));
  } catch {
    return [];
  }
}

/**
 * Descifra una fuente .odttf del Word (ECMA-376, 17.8.1; lo mismo que hace docx-preview): los primeros 32 bytes van
 * mezclados (XOR) con la clave GUID leída al revés. Devuelve una copia; no toca el original.
 */
export function deobfuscateFont(data: Uint8Array, guid: string): Uint8Array {
  const hex = guid.replace(/[{}-]/g, "");
  const out = new Uint8Array(data);
  if (!/^[0-9a-f]{32}$/i.test(hex)) return out;
  const key = new Array<number>(16);
  for (let i = 0; i < 16; i++) key[15 - i] = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  for (let i = 0; i < 32 && i < out.length; i++) out[i] = out[i]! ^ key[i % 16]!;
  return out;
}

function sfntTag(bytes: Uint8Array, at = 0): string {
  return String.fromCharCode(bytes[at] ?? 0, bytes[at + 1] ?? 0, bytes[at + 2] ?? 0, bytes[at + 3] ?? 0);
}

function isSfnt(bytes: Uint8Array): boolean {
  const tag = sfntTag(bytes);
  return tag === "\0\x01\0\0" || tag === "OTTO" || tag === "true" || tag === "ttcf";
}

/** OS/2 fsType: 2 = licencia restringida, 0x200 = solo mapas de bits. */
function restrictedLicense(bytes: Uint8Array): boolean {
  if (bytes.length < 16) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const base = sfntTag(bytes) === "ttcf" ? view.getUint32(12) : 0;
  if (base + 12 > bytes.length) return false;
  const count = view.getUint16(base + 4);
  for (let i = 0; i < count; i++) {
    const rec = base + 12 + i * 16;
    if (rec + 16 > bytes.length) break;
    if (sfntTag(bytes, rec) !== "OS/2") continue;
    const offset = view.getUint32(rec + 8);
    if (offset + 10 > bytes.length) return false;
    const fsType = view.getUint16(offset + 8);
    return (fsType & 0x000f) === 0x0002 || (fsType & 0x0200) !== 0;
  }
  return false;
}

function styleName(weight: number, style: string): string {
  const parts = [weight >= 600 ? "negrita" : "", style === "normal" ? "" : "cursiva"].filter(Boolean);
  return parts.length ? parts.join(" ") : "normal";
}

/** Las fuentes incrustadas del Word ya descifradas, una por cara, y avisos si alguna no se puede usar. */
export async function embeddedFonts(bytes: Uint8Array): Promise<{ faces: EmbeddedFace[]; warnings: string[] }> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return { faces: [], warnings: [] };
  }
  const faces: EmbeddedFace[] = [];
  const warnings: string[] = [];
  for (const ref of await refsFromZip(zip)) {
    const raw = await zip.file(ref.path)?.async("uint8array");
    const name = `«${ref.family}» (${styleName(ref.weight, ref.style)})`;
    if (!raw) {
      warnings.push(`Falta el archivo de la fuente incrustada ${name} dentro del Word: se usa la sustituta o hay que subirla.`);
      continue;
    }
    // A veces el Word guarda la fuente sin cifrar: se usa tal cual.
    const font = isSfnt(raw) ? new Uint8Array(raw) : deobfuscateFont(raw, ref.fontKey);
    if (!isSfnt(font)) {
      warnings.push(`La fuente incrustada ${name} no se pudo leer: se usa la sustituta o hay que subirla.`);
      continue;
    }
    const restricted = restrictedLicense(font);
    if (restricted) {
      warnings.push(
        `La licencia de la fuente incrustada ${name} no permite meterla en un PDF: conviene subir una versión con licencia.`,
      );
    }
    faces.push({ ...ref, bytes: font, restricted });
  }
  return { faces, warnings };
}

/**
 * Claves faceKey de las caras incrustadas enteras: pasarlas como `skip` para que la sustituta no tape a la del Word.
 * Una cara recortada (`subsetted`: solo las letras que usa el texto) no se salta: la sustituta queda detrás para las
 * letras que agrega Pliego (números de página, «‑», mayúsculas de las versalitas). Ver subsetFallbacks.
 */
export function embeddedSkips(
  faces: Array<{ family: string; weight: number; style: string; subsetted?: boolean }>,
): Set<string> {
  return new Set(faces.filter((face) => !face.subsetted).map((face) => faceKey(face.family, face.weight, face.style)));
}

/** Claves faceKey de las caras recortadas del Word que tienen sustituta: para installSubstitutes y addSubstituteFallback. */
export function subsetFallbacks(
  faces: Array<{ family: string; weight: number; style: string; subsetted?: boolean }>,
): Set<string> {
  const known = new Set(SUBSTITUTE_FACES.map((face) => faceKey(face.family, face.weight, face.style)));
  return new Set(
    faces
      .filter((face) => face.subsetted)
      .map((face) => faceKey(face.family, face.weight, face.style))
      .filter((key) => known.has(key)),
  );
}

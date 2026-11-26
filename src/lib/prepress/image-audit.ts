import JSZip from "jszip";

/**
 * Revisión de las imágenes y los dibujos del Word limpio, antes de armar el PDF. Solo lee: no recomprime, no convierte
 * y no toca el .docx. Todo lo que el PDF no puede llevar tal cual vuelve como aviso para quien edita.
 */

export type ImageFormat = "png" | "jpeg" | "gif" | "bmp" | "webp" | "tiff" | "emf" | "wmf" | "svg" | "comprimido" | "desconocido";

const ascii = (bytes: Uint8Array, start: number, length: number) => {
  let text = "";
  for (let i = start; i < start + length && i < bytes.length; i++) text += String.fromCharCode(bytes[i]!);
  return text;
};
const u16 = (b: Uint8Array, p: number, le = false) => (le ? (b[p]! | (b[p + 1]! << 8)) : ((b[p]! << 8) | b[p + 1]!));
const u32 = (b: Uint8Array, p: number, le = false) =>
  le
    ? (b[p]! | (b[p + 1]! << 8) | (b[p + 2]! << 16) | (b[p + 3]! << 24)) >>> 0
    : ((b[p]! << 24) | (b[p + 1]! << 16) | (b[p + 2]! << 8) | b[p + 3]!) >>> 0;

/** Formato por los primeros bytes (el nombre del archivo puede mentir). */
export function sniffImage(bytes: Uint8Array): ImageFormat {
  if (bytes.length < 4) return "desconocido";
  if (bytes[0] === 0x89 && ascii(bytes, 1, 3) === "PNG") return "png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "jpeg";
  if (ascii(bytes, 0, 4) === "GIF8") return "gif";
  if (ascii(bytes, 0, 4) === "RIFF" && ascii(bytes, 8, 4) === "WEBP") return "webp";
  if (ascii(bytes, 0, 4) === "II*\0" || ascii(bytes, 0, 4) === "MM\0*") return "tiff";
  if (u32(bytes, 0, true) === 1 && ascii(bytes, 40, 4) === " EMF") return "emf";
  if (u32(bytes, 0, true) === 0x9ac6cdd7) return "wmf";
  if ((bytes[0] === 1 || bytes[0] === 2) && bytes[1] === 0 && bytes[2] === 9 && bytes[3] === 0) return "wmf";
  if (bytes[0] === 0x1f && bytes[1] === 0x8b) return "comprimido";
  if (ascii(bytes, 0, 2) === "BM") return "bmp";
  const head = new TextDecoder("utf-8", { fatal: false }).decode(bytes.subarray(0, 512)).replace(/^﻿/, "").trimStart();
  if (/^(<\?xml[\s\S]*?)?(<!--[\s\S]*?-->\s*)*(<!DOCTYPE svg[^>]*>\s*)?<svg\b/i.test(head)) return "svg";
  return "desconocido";
}

export type PngInfo = {
  width: number;
  height: number;
  bitDepth: number;
  /** 0 gris, 2 RGB, 3 paleta, 4 gris + alfa, 6 RGB + alfa. */
  colorType: number;
  /** Canal alfa o tRNS: puede tener partes transparentes. */
  alpha: boolean;
  /** Marca sRGB explícita. */
  srgb: boolean;
  iccName: string | null;
  /** Perfil ICC todavía comprimido (zlib), tal como viene en iCCP. */
  iccData: Uint8Array | null;
};

export function readPng(bytes: Uint8Array): PngInfo | null {
  if (sniffImage(bytes) !== "png" || bytes.length < 33 || ascii(bytes, 12, 4) !== "IHDR") return null;
  const info: PngInfo = {
    width: u32(bytes, 16),
    height: u32(bytes, 20),
    bitDepth: bytes[24]!,
    colorType: bytes[25]!,
    alpha: bytes[25] === 4 || bytes[25] === 6,
    srgb: false,
    iccName: null,
    iccData: null,
  };
  let p = 8;
  while (p + 12 <= bytes.length) {
    const length = u32(bytes, p);
    const type = ascii(bytes, p + 4, 4);
    const start = p + 8;
    const end = start + length;
    if (end > bytes.length || type === "IDAT" || type === "IEND") break;
    if (type === "tRNS") info.alpha = true;
    else if (type === "sRGB") info.srgb = true;
    else if (type === "iCCP") {
      let zero = start;
      while (zero < end && bytes[zero] !== 0) zero++;
      info.iccName = new TextDecoder("latin1").decode(bytes.subarray(start, zero));
      // Después del nombre: un byte de método de compresión y el perfil en zlib.
      info.iccData = bytes.subarray(Math.min(zero + 2, end), end);
    }
    p = end + 4;
  }
  return info;
}

export type JpegInfo = {
  width: number;
  height: number;
  /** 1 gris, 3 RGB/YCbCr, 4 CMYK/YCCK. */
  components: number;
  precision: number;
  /** EXIF Orientation (1 = derecha). */
  orientation: number;
  /** Marcador Adobe APP14: dice cómo están guardados los CMYK. */
  adobe: boolean;
  icc: Uint8Array | null;
};

function exifOrientation(b: Uint8Array, tiff: number, end: number): number {
  if (tiff + 8 > end) return 1;
  const le = ascii(b, tiff, 2) === "II";
  const ifd = tiff + u32(b, tiff + 4, le);
  if (ifd + 2 > end) return 1;
  const count = u16(b, ifd, le);
  for (let i = 0; i < count; i++) {
    const entry = ifd + 2 + i * 12;
    if (entry + 12 > end) break;
    if (u16(b, entry, le) === 0x0112) return u16(b, entry + 8, le) || 1;
  }
  return 1;
}

export function readJpeg(bytes: Uint8Array): JpegInfo | null {
  if (sniffImage(bytes) !== "jpeg") return null;
  const info: JpegInfo = { width: 0, height: 0, components: 0, precision: 8, orientation: 1, adobe: false, icc: null };
  const icc: Array<[number, Uint8Array]> = [];
  let p = 2;
  while (p + 4 <= bytes.length) {
    if (bytes[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = bytes[p + 1]!;
    if (marker === 0xff) {
      p++;
      continue;
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      p += 2;
      continue;
    }
    if (marker === 0xd9 || marker === 0xda) break;
    const length = u16(bytes, p + 2);
    const start = p + 4;
    const end = p + 2 + length;
    if (length < 2 || end > bytes.length) break;
    if (marker === 0xe1 && ascii(bytes, start, 6) === "Exif\0\0") info.orientation = exifOrientation(bytes, start + 6, end);
    else if (marker === 0xe2 && ascii(bytes, start, 12) === "ICC_PROFILE\0") icc.push([bytes[start + 12] ?? 0, bytes.subarray(start + 14, end)]);
    else if (marker === 0xee && ascii(bytes, start, 5) === "Adobe") info.adobe = true;
    else if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      info.precision = bytes[start]!;
      info.height = u16(bytes, start + 1);
      info.width = u16(bytes, start + 3);
      info.components = bytes[start + 5]!;
    }
    p = end;
  }
  if (icc.length) {
    icc.sort((a, b) => a[0] - b[0]);
    const total = icc.reduce((sum, [, part]) => sum + part.length, 0);
    const joined = new Uint8Array(total);
    let at = 0;
    for (const [, part] of icc) {
      joined.set(part, at);
      at += part.length;
    }
    info.icc = joined;
  }
  return info;
}

/** Espacio de color y nombre de un perfil ICC (etiqueta desc, v2 o v4). */
export function iccInfo(profile: Uint8Array): { space: string; description: string } | null {
  if (profile.length < 132 || ascii(profile, 36, 4) !== "acsp") return null;
  const space = ascii(profile, 16, 4).trim();
  const count = u32(profile, 128);
  let description = "";
  for (let i = 0; i < count && 132 + i * 12 + 12 <= profile.length; i++) {
    const rec = 132 + i * 12;
    if (ascii(profile, rec, 4) !== "desc") continue;
    const offset = u32(profile, rec + 4);
    const size = u32(profile, rec + 8);
    if (offset + 12 > profile.length) break;
    const kind = ascii(profile, offset, 4);
    if (kind === "desc") {
      const length = u32(profile, offset + 8);
      description = new TextDecoder("latin1").decode(profile.subarray(offset + 12, Math.min(offset + 12 + length, offset + size, profile.length)));
    } else if (kind === "mluc" && offset + 28 <= profile.length) {
      const length = u32(profile, offset + 20);
      const at = offset + u32(profile, offset + 24);
      let text = "";
      for (let j = at; j + 1 < at + length && j + 1 < profile.length; j += 2) text += String.fromCharCode(u16(profile, j));
      description = text;
    }
    break;
  }
  return { space, description: description.replace(/\0/g, "").trim() };
}

export function isSrgb(description: string): boolean {
  return /\bs\s?rgb|iec\s*61966-2/i.test(description);
}

async function inflate(data: Uint8Array): Promise<Uint8Array | null> {
  if (typeof DecompressionStream === "undefined") return null;
  try {
    const stream = new Blob([new Uint8Array(data)]).stream().pipeThrough(new DecompressionStream("deflate"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch {
    return null;
  }
}

/** Perfil de color de una PNG o una JPEG: null si no trae, o si es sRGB. */
async function foreignProfile(format: ImageFormat, bytes: Uint8Array): Promise<string | null> {
  if (format === "png") {
    const png = readPng(bytes);
    if (!png || png.srgb || !png.iccData) return null;
    const profile = await inflate(png.iccData);
    const info = profile ? iccInfo(profile) : null;
    const name = info?.description || png.iccName || "";
    if (info && info.space !== "RGB") return null;
    return name && isSrgb(name) ? null : name || "sin nombre";
  }
  if (format === "jpeg") {
    const jpeg = readJpeg(bytes);
    if (!jpeg?.icc) return null;
    const info = iccInfo(jpeg.icc);
    if (!info || info.space !== "RGB") return null;
    return isSrgb(info.description) ? null : info.description || "sin nombre";
  }
  return null;
}

/* ---------- Avisos del .docx ---------- */

type Issue = { one: string; many: string };

const IMAGE_ISSUES = {
  vector: {
    one: "La imagen {lista} está en EMF o WMF (dibujo de Office): no va a salir en el PDF. Reemplazala por PNG o JPG en el Word.",
    many: "Las imágenes {lista} están en EMF o WMF (dibujo de Office): no van a salir en el PDF. Reemplazalas por PNG o JPG en el Word.",
  },
  tiff: {
    one: "La imagen {lista} está en TIFF: no va a salir en el PDF. Reemplazala por PNG o JPG en el Word.",
    many: "Las imágenes {lista} están en TIFF: no van a salir en el PDF. Reemplazalas por PNG o JPG en el Word.",
  },
  unknown: {
    one: "La imagen {lista} está en un formato que la app no reconoce: puede no salir en el PDF. Reemplazala por PNG o JPG en el Word.",
    many: "Las imágenes {lista} están en un formato que la app no reconoce: pueden no salir en el PDF. Reemplazalas por PNG o JPG en el Word.",
  },
  redraw: {
    one: "La imagen {lista} está en GIF, BMP, WebP o SVG: al pasar al PDF se redibuja. Para que vaya tal cual, reemplazala por PNG o JPG en el Word.",
    many: "Las imágenes {lista} están en GIF, BMP, WebP o SVG: al pasar al PDF se redibujan. Para que vayan tal cual, reemplazalas por PNG o JPG en el Word.",
  },
  svg: {
    one: "La imagen {lista} es un SVG: el PDF usa la copia PNG que guardó Word, que puede tener poca resolución. Revisá sus ppp abajo o insertala como PNG grande.",
    many: "Las imágenes {lista} son SVG: el PDF usa las copias PNG que guardó Word, que pueden tener poca resolución. Revisá sus ppp abajo o insertalas como PNG grandes.",
  },
  linked: {
    one: "La imagen {lista} está vinculada, no metida en el Word: no va a salir. En Word, Archivo → Información → Editar vínculos → «Guardar imagen en el documento».",
    many: "Las imágenes {lista} están vinculadas, no metidas en el Word: no van a salir. En Word, Archivo → Información → Editar vínculos → «Guardar imagen en el documento».",
  },
  missing: {
    one: "Falta el archivo de la imagen {lista} dentro del Word: no va a salir. Volvé a insertarla.",
    many: "Faltan los archivos de las imágenes {lista} dentro del Word: no van a salir. Volvé a insertarlas.",
  },
  orientation: {
    one: "La foto {lista} trae el giro guardado aparte (EXIF): en el PDF puede salir girada. Girá el archivo de verdad (y guardalo) antes de insertarlo.",
    many: "Las fotos {lista} traen el giro guardado aparte (EXIF): en el PDF pueden salir giradas. Girá los archivos de verdad (y guardalos) antes de insertarlos.",
  },
  profile: {
    one: "La imagen {lista} tiene un perfil de color que no es sRGB: en el PDF los colores pueden cambiar. Convertila a sRGB antes de insertarla.",
    many: "Las imágenes {lista} tienen un perfil de color que no es sRGB: en el PDF los colores pueden cambiar. Convertilas a sRGB antes de insertarlas.",
  },
  cmyk: {
    one: "La imagen {lista} es un JPEG en CMYK sin la marca de Adobe: en el PDF puede salir en negativo. Volvé a guardarla desde Photoshop o pasala a RGB.",
    many: "Las imágenes {lista} son JPEG en CMYK sin la marca de Adobe: en el PDF pueden salir en negativo. Volvé a guardarlas desde Photoshop o pasalas a RGB.",
  },
  precision: {
    one: "La imagen {lista} es un JPEG de 12 bits: el PDF no lo admite y puede no salir. Guardala como JPEG común (8 bits).",
    many: "Las imágenes {lista} son JPEG de 12 bits: el PDF no los admite y pueden no salir. Guardalas como JPEG común (8 bits).",
  },
  deep: {
    one: "La imagen {lista} es una PNG de 16 bits por canal: en el PDF pasa a 8 bits (en papel no se nota).",
    many: "Las imágenes {lista} son PNG de 16 bits por canal: en el PDF pasan a 8 bits (en papel no se nota).",
  },
  alpha: {
    one: "La imagen {lista} tiene transparencia: en el PDF para imprenta (y en B/N) se compone sobre blanco. Si debajo hay color o texto, guardala sin transparencia.",
    many: "Las imágenes {lista} tienen transparencia: en el PDF para imprenta (y en B/N) se componen sobre blanco. Si debajo hay color o texto, guardalas sin transparencia.",
  },
} satisfies Record<string, Issue>;

type IssueKey = keyof typeof IMAGE_ISSUES;

function list(names: string[]): string {
  const unique = [...new Set(names)];
  const shown = unique.slice(0, 5).map((name) => `«${name}»`);
  return unique.length > 5 ? `${shown.join(", ")} y ${unique.length - 5} más` : shown.join(", ");
}

function phrase(issue: Issue, names: string[]): string {
  const count = new Set(names).size;
  return (count === 1 ? issue.one : issue.many).replace("{lista}", list(names));
}

const USED_PART = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

function decodeXml(text: string): string {
  return text
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function attr(tag: string, name: string): string | undefined {
  const value = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];
  return value === undefined ? undefined : decodeXml(value);
}

export function relsPathOf(part: string): string {
  const slash = part.lastIndexOf("/");
  return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
}

export function resolveTarget(part: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = part.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "..") parts.pop();
    else if (segment && segment !== ".") parts.push(segment);
  }
  return parts.join("/");
}

export type Relationship = { target: string; type: string; external: boolean };

export function parseRels(xml: string): Map<string, Relationship> {
  const rels = new Map<string, Relationship>();
  for (const tag of xml.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = attr(tag, "Id");
    const target = attr(tag, "Target");
    if (!id || target === undefined) continue;
    rels.set(id, { target, type: attr(tag, "Type") ?? "", external: attr(tag, "TargetMode") === "External" });
  }
  return rels;
}

const EFFECTS: Array<[RegExp, string]> = [
  [/<a:grayscl\b/, "escala de grises"],
  [/<a:lum\b/, "brillo o contraste"],
  [/<a:clrChange\b/, "color transparente"],
  [/<a:duotone\b/, "recoloreado"],
  [/<a:biLevel\b/, "blanco y negro"],
  [/<a:alphaModFix\b/, "transparencia"],
  [/<a14:imgEffect\b|<a14:imgProps\b/, "efecto artístico"],
  [/\sgrayscale="(?:t|true|1)"/, "escala de grises"],
  [/\sbilevel="(?:t|true|1)"/, "blanco y negro"],
  [/\s(?:gain|blacklevel)="/, "brillo o contraste"],
];

/** Lee el Word limpio y devuelve avisos sobre imágenes, gráficos y formas que el PDF no puede llevar tal cual. */
export async function auditDocx(bytes: Uint8Array): Promise<string[]> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return [];
  }
  const found = new Map<IssueKey, string[]>();
  const add = (key: IssueKey, name: string) => found.set(key, [...(found.get(key) ?? []), name]);

  const media = new Map<string, { svg: boolean }>();
  let charts = 0;
  let smartArt = 0;
  let shapes = 0;
  let textBoxes = 0;
  let objects = 0;
  let vml = 0;
  let cropped = 0;
  let rotated = 0;
  const effects = new Map<string, number>();

  for (const part of Object.keys(zip.files).filter((path) => USED_PART.test(path)).sort()) {
    const raw = await zip.file(part)?.async("string");
    if (!raw) continue;
    const relsXml = await zip.file(relsPathOf(part))?.async("string");
    const rels = relsXml ? parseRels(relsXml) : new Map<string, Relationship>();

    // docx-preview siempre dibuja el mc:Fallback; para contar sin repetir, se mira solo la versión principal.
    const xml = raw.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, "");

    const refer = (id: string | undefined, svg: boolean, linkAttr: boolean) => {
      if (!id) return;
      const rel = rels.get(id);
      if (!rel) return;
      const name = rel.target.split("/").pop() || rel.target;
      if (rel.external || linkAttr) {
        add("linked", name);
        return;
      }
      const path = resolveTarget(part, rel.target);
      const known = media.get(path);
      media.set(path, { svg: svg || Boolean(known?.svg) });
    };
    for (const tag of raw.match(/<a:blip\b[^>]*>/g) ?? []) {
      refer(attr(tag, "r:embed"), false, false);
      refer(attr(tag, "r:link"), false, true);
    }
    for (const tag of raw.match(/<[\w]+:svgBlip\b[^>]*>/g) ?? []) refer(attr(tag, "r:embed"), true, false);
    for (const tag of raw.match(/<v:imagedata\b[^>]*>/g) ?? []) refer(attr(tag, "r:id") ?? attr(tag, "o:relid"), false, false);

    const objectBlocks = xml.match(/<w:object\b[\s\S]*?<\/w:object>/g) ?? [];
    objects += objectBlocks.length;
    const rest = xml.replace(/<w:object\b[\s\S]*?<\/w:object>/g, "");

    for (const drawing of rest.match(/<w:drawing\b[\s\S]*?<\/w:drawing>/g) ?? []) {
      if (/<c:chart\b|<cx:chart\b|drawingml\/2006\/chart"|drawingml\/2014\/chartex"/.test(drawing)) {
        charts++;
        continue;
      }
      if (/<dgm:relIds\b/.test(drawing)) {
        smartArt++;
        continue;
      }
      if (/<wps:wsp\b|<wpg:wgp\b|<wpc:wpc\b/.test(drawing)) {
        shapes++;
        if (/<w:txbxContent\b|<wps:txbx\b/.test(drawing)) textBoxes++;
        continue;
      }
      if (!/<pic:pic\b/.test(drawing)) continue;
      for (const [pattern, label] of EFFECTS) if (pattern.test(drawing)) effects.set(label, (effects.get(label) ?? 0) + 1);
      const xfrm = /<pic:spPr\b[\s\S]*?<a:xfrm\b([^>]*)>/.exec(drawing)?.[1] ?? "";
      if (/\sflip[HV]="(?:1|true)"/.test(xfrm)) effects.set("volteo (espejo)", (effects.get("volteo (espejo)") ?? 0) + 1);
      const rot = Number(attr(xfrm, "rot") ?? 0);
      if (rot % 21600000 !== 0) rotated++;
      const rect = /<a:srcRect\b[^>]*>/.exec(drawing)?.[0];
      if (rect && ["l", "t", "r", "b"].some((side) => Number(attr(rect, side) ?? 0) !== 0)) cropped++;
    }

    for (const pict of rest.match(/<w:pict\b[\s\S]*?<\/w:pict>/g) ?? []) {
      if (/<v:imagedata\b/.test(pict) && !/<v:textbox\b/.test(pict)) {
        vml++;
        for (const [pattern, label] of EFFECTS) if (pattern.test(pict)) effects.set(label, (effects.get(label) ?? 0) + 1);
        if (/\scrop(?:left|top|right|bottom)="/.test(pict)) cropped++;
        continue;
      }
      if (/<v:(?:shape|rect|roundrect|oval|line|polyline|arc|group|textbox)\b/.test(pict)) {
        shapes++;
        if (/<w:txbxContent\b|<v:textbox\b/.test(pict)) textBoxes++;
      }
    }
  }

  if (!charts) {
    charts = Object.keys(zip.files).filter((path) => /^word\/charts\/chart[^/]*\.xml$/.test(path)).length;
  }
  if (!smartArt) {
    smartArt = Object.keys(zip.files).filter((path) => /^word\/diagrams\/data[^/]*\.xml$/.test(path)).length;
  }

  for (const [path, { svg }] of [...media].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))) {
    const name = path.split("/").pop() || path;
    const data = await zip.file(path)?.async("uint8array");
    if (!data) {
      add("missing", name);
      continue;
    }
    const format = sniffImage(data);
    if (format === "emf" || format === "wmf" || format === "comprimido") add("vector", name);
    else if (format === "tiff") add("tiff", name);
    else if (format === "gif" || format === "bmp" || format === "webp") add("redraw", name);
    else if (format === "svg") {
      if (!svg) add("redraw", name);
      else add("svg", name);
    } else if (format === "desconocido") add("unknown", name);
    else if (format === "jpeg") {
      const jpeg = readJpeg(data);
      if (jpeg && jpeg.orientation !== 1) add("orientation", name);
      if (jpeg && jpeg.components === 4 && !jpeg.adobe) add("cmyk", name);
      if (jpeg && jpeg.precision > 8) add("precision", name);
    } else if (format === "png") {
      const png = readPng(data);
      if (png?.bitDepth === 16) add("deep", name);
      if (png?.alpha) add("alpha", name);
    }
    const profile = await foreignProfile(format, data);
    if (profile) add("profile", profile === "sin nombre" ? name : `${name}» («${profile}`);
  }

  const warnings: string[] = [];
  for (const key of Object.keys(IMAGE_ISSUES) as IssueKey[]) {
    const names = found.get(key);
    if (names?.length) warnings.push(phrase(IMAGE_ISSUES[key], names));
  }
  if (charts) {
    warnings.push(
      charts === 1
        ? "Hay un gráfico de Word (de los que se editan con datos): no se puede pasar al PDF. Copialo y pegalo como imagen (Pegado especial → Imagen PNG)."
        : `Hay ${charts} gráficos de Word (de los que se editan con datos): no se pueden pasar al PDF. Copialos y pegalos como imagen (Pegado especial → Imagen PNG).`,
    );
  }
  if (smartArt) {
    warnings.push(
      smartArt === 1
        ? "Hay un SmartArt: no se puede pasar al PDF. Copialo y pegalo como imagen (Pegado especial → Imagen PNG)."
        : `Hay ${smartArt} SmartArt: no se pueden pasar al PDF. Copialos y pegalos como imagen (Pegado especial → Imagen PNG).`,
    );
  }
  if (objects) {
    warnings.push(
      objects === 1
        ? "Hay un objeto incrustado (una ecuación antigua, una planilla u otro programa): no va a salir en el PDF. Pegalo como imagen."
        : `Hay ${objects} objetos incrustados (ecuaciones antiguas, planillas u otros programas): no van a salir en el PDF. Pegalos como imagen.`,
    );
  }
  if (shapes) {
    const what = textBoxes ? (shapes === 1 ? "un cuadro de texto o forma" : `${shapes} cuadros de texto o formas`) : shapes === 1 ? "una forma (autoforma)" : `${shapes} formas (autoformas)`;
    warnings.push(
      shapes === 1
        ? `Hay ${what}: puede no salir en el PDF o salir como imagen. Revisalo en la vista previa; si falla, pasá el texto al cuerpo o pegá la forma como imagen.`
        : `Hay ${what}: pueden no salir en el PDF o salir como imagen. Revisalos en la vista previa; si fallan, pasá el texto al cuerpo o pegá las formas como imagen.`,
    );
  }
  if (vml) {
    warnings.push(
      vml === 1
        ? "Hay una imagen guardada en el formato viejo de Word: al pasar al PDF puede no salir o perder resolución. Volvé a insertarla con Insertar → Imágenes."
        : `Hay ${vml} imágenes guardadas en el formato viejo de Word: al pasar al PDF pueden no salir o perder resolución. Volvé a insertarlas con Insertar → Imágenes.`,
    );
  }
  if (effects.size) {
    const total = Math.max(...effects.values());
    const labels = [...effects.keys()].join(", ");
    warnings.push(
      total === 1 && effects.size === 1
        ? `Una imagen tiene un efecto de Word (${labels}): el efecto no se aplica en el PDF. Aplicalo a la imagen antes de insertarla.`
        : `Hay imágenes con efectos de Word (${labels}): esos efectos no se aplican en el PDF. Aplicalos a las imágenes antes de insertarlas.`,
    );
  }
  if (cropped) {
    warnings.push(
      cropped === 1
        ? "Una imagen está recortada dentro del Word: en el PDF puede salir sin el recorte. Recortá el archivo antes de insertarlo."
        : `Hay ${cropped} imágenes recortadas dentro del Word: en el PDF pueden salir sin el recorte. Recortá los archivos antes de insertarlos.`,
    );
  }
  if (rotated) {
    warnings.push(
      rotated === 1
        ? "Una imagen está girada dentro del Word: en el PDF puede salir deformada. Girá el archivo antes de insertarlo."
        : `Hay ${rotated} imágenes giradas dentro del Word: en el PDF pueden salir deformadas. Girá los archivos antes de insertarlos.`,
    );
  }
  return warnings;
}

/* ---------- Resolución de cada imagen ---------- */

export type ImageLevel = "ok" | "baja" | "linea" | "rechazo" | "sin-medida" | "rota";

/** Recorte de Word (a:srcRect), en milésimas de porcentaje: 25000 = 25 %. */
export type SrcRect = { l?: number; t?: number; r?: number; b?: number };

export type LevelInput = {
  naturalWidth: number;
  naturalHeight: number;
  /** Tamaño en la página, en pulgadas. */
  widthIn: number;
  heightIn: number;
  crop?: SrcRect;
  /** Dibujo de línea (1 bit): la imprenta pide 600 ppp. */
  lineArt?: boolean;
};

export type LevelResult = { ppi: number; level: ImageLevel; note: string; visibleWidth: number; visibleHeight: number };

/** Píxeles que se ven después del recorte. */
export function visiblePixels(natural: number, start = 0, end = 0): number {
  const fraction = 1 - (start + end) / 100000;
  return Math.max(0, Math.round(natural * Math.min(1, fraction)));
}

/** ppp de la parte visible y nivel para imprenta (Lightning Source: 300 ppp, 600 para dibujo de línea, menos de 72 se rechaza). */
export function imageLevels(input: LevelInput): LevelResult {
  const crop = input.crop ?? {};
  const visibleWidth = visiblePixels(input.naturalWidth, crop.l, crop.r);
  const visibleHeight = visiblePixels(input.naturalHeight, crop.t, crop.b);
  if (!input.naturalWidth || !input.naturalHeight) {
    return { ppi: 0, level: "rota", note: "no se pudo mostrar: no va a salir", visibleWidth, visibleHeight };
  }
  if (!(input.widthIn > 0) || !(input.heightIn > 0)) {
    return { ppi: 0, level: "sin-medida", note: "sin medida", visibleWidth, visibleHeight };
  }
  const ppi = Math.round(Math.min(visibleWidth / input.widthIn, visibleHeight / input.heightIn));
  if (ppi < 72) return { ppi, level: "rechazo", note: "rechazo seguro: menos de 72 ppp", visibleWidth, visibleHeight };
  if (input.lineArt && ppi < 600) return { ppi, level: "linea", note: "baja para dibujo de línea (pide 600 ppp)", visibleWidth, visibleHeight };
  if (ppi < 300) return { ppi, level: "baja", note: "baja para imprenta (pide 300 ppp)", visibleWidth, visibleHeight };
  return { ppi, level: "ok", note: "sin recomprimir", visibleWidth, visibleHeight };
}

/** docx-preview dibuja el recorte con scale(1/(1-l-r), 1/(1-t-b)): de ahí sale cuánto se ve. */
export function cropFromTransform(transform: string): SrcRect {
  const match = /scale\(\s*([\d.e+-]+)(?:\s*,\s*([\d.e+-]+))?\s*\)/i.exec(transform);
  if (!match) return {};
  const sx = Number(match[1]);
  const sy = Number(match[2] ?? match[1]);
  const cut = (scale: number) => (Number.isFinite(scale) && scale > 1 ? (1 - 1 / scale) * 100000 : 0);
  return { l: cut(sx), t: cut(sy) };
}

export type RenderedImage = {
  label: string;
  /** Página del visor (1 = primera). */
  page: number;
  pixels: string;
  ppi: number;
  level: ImageLevel;
  note: string;
};

function isBitonal(bytes: Uint8Array): boolean {
  const png = readPng(bytes);
  return Boolean(png && png.bitDepth === 1 && (png.colorType === 0 || png.colorType === 3));
}

async function sourceBytes(url: string): Promise<Uint8Array | null> {
  if (!url) return null;
  try {
    const response = await fetch(url);
    return response.ok ? new Uint8Array(await response.arrayBuffer()) : null;
  } catch {
    return null;
  }
}

function naturalSize(url: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const probe = new Image();
    probe.onload = () => resolve({ width: probe.naturalWidth, height: probe.naturalHeight });
    probe.onerror = () => resolve({ width: 0, height: 0 });
    probe.src = url;
  });
}

/** Mide cada imagen ya compuesta (solo lee el DOM; usar con la página a escala 1 o con medidas de maqueta). */
export async function auditRenderedImages(root: ParentNode): Promise<RenderedImage[]> {
  const sections = [...root.querySelectorAll("section.docx")];
  const pageOf = (node: Element) => {
    const section = node.closest("section.docx");
    return section ? sections.indexOf(section) + 1 : 0;
  };
  const results: RenderedImage[] = [];
  const images = [...root.querySelectorAll("section.docx img")] as HTMLImageElement[];
  for (const [index, img] of images.entries()) {
    const label = img.alt || `Imagen ${index + 1}`;
    const page = pageOf(img);
    if (!img.naturalWidth || !img.naturalHeight) {
      results.push({ label, page, pixels: "—", ppi: 0, level: "rota", note: "no se pudo mostrar: no va a salir" });
      continue;
    }
    const bytes = await sourceBytes(img.currentSrc || img.src);
    const format = bytes ? sniffImage(bytes) : "desconocido";
    const crop = cropFromTransform(img.style.transform || "");
    const result = imageLevels({
      naturalWidth: img.naturalWidth,
      naturalHeight: img.naturalHeight,
      widthIn: img.clientWidth / 96,
      heightIn: img.clientHeight / 96,
      crop,
      lineArt: bytes ? isBitonal(bytes) : false,
    });
    const notes = [result.note];
    if (bytes && format !== "png" && format !== "jpeg") notes.push("se redibuja");
    if ((crop.l ?? 0) > 0 || (crop.t ?? 0) > 0) notes.push("recortada: revisá el recorte en el PDF");
    if (/rotate\(/.test(img.style.transform || "")) notes.push("girada: revisala en el PDF");
    const cropped = result.visibleWidth !== img.naturalWidth || result.visibleHeight !== img.naturalHeight;
    results.push({
      label,
      page,
      pixels: cropped
        ? `${result.visibleWidth}×${result.visibleHeight} de ${img.naturalWidth}×${img.naturalHeight}`
        : `${img.naturalWidth}×${img.naturalHeight}`,
      ppi: result.ppi,
      level: result.level,
      note: notes.join(" · "),
    });
  }
  // Imágenes del formato viejo de Word (VML): el visor las pone dentro de un <svg>.
  const legacy = [...root.querySelectorAll("section.docx svg image")] as SVGImageElement[];
  for (const [index, image] of legacy.entries()) {
    const url = image.getAttribute("href") || image.getAttribute("xlink:href") || "";
    const label = `Imagen antigua ${index + 1}`;
    const page = pageOf(image);
    const size = url ? await naturalSize(url) : { width: 0, height: 0 };
    let box = { width: 0, height: 0 };
    try {
      box = image.getBBox();
    } catch {
      /* Sin caja: queda sin medida. */
    }
    const result = imageLevels({
      naturalWidth: size.width,
      naturalHeight: size.height,
      widthIn: box.width / 96,
      heightIn: box.height / 96,
    });
    results.push({
      label,
      page,
      pixels: size.width ? `${size.width}×${size.height}` : "—",
      ppi: result.ppi,
      level: result.level,
      note:
        result.level === "rota"
          ? result.note
          : `${result.note} · formato viejo de Word: puede no salir en el PDF; volvé a insertarla`,
    });
  }
  return results;
}

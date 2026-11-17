import { parseFontFaces, type ParsedFace } from "@/lib/prepress/font-face";

export type ShelfFace = ParsedFace & {
  id: string;
  bytes: Uint8Array;
  fileName: string;
};

const DB_NAME = "pliego-fonts";
const STORE = "faces";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(STORE, { keyPath: "id" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function loadShelf(): Promise<ShelfFace[]> {
  try {
    const db = await openDb();
    const faces = await new Promise<ShelfFace[]>((resolve, reject) => {
      const request = db.transaction(STORE).objectStore(STORE).getAll();
      request.onsuccess = () => resolve(request.result as ShelfFace[]);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return faces.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes) }));
  } catch {
    return [];
  }
}

async function persist(face: ShelfFace) {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).put({
      ...face,
      bytes: face.bytes.buffer.slice(face.bytes.byteOffset, face.bytes.byteOffset + face.bytes.byteLength),
    });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
  db.close();
}

export async function rememberFace(face: ShelfFace) {
  try {
    await persist(face);
  } catch {
    /* Si el navegador no guarda, la fuente sigue viva en esta sesión. */
  }
}

export async function forgetFace(id: string) {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
  db.close();
}

export async function readFontFile(file: File): Promise<ShelfFace[]> {
  if (file.size > 32 * 1024 * 1024) throw new Error(`${file.name} pasa de 32 MB.`);
  const bytes = new Uint8Array(await file.arrayBuffer());
  return parseFontFaces(bytes).map((face, index) => ({
    ...face,
    bytes,
    fileName: file.name,
    id: `${file.name}:${index}:${face.family}:${face.weight}:${face.style}`,
  }));
}

const STYLE_ID = "pliego-uploaded-fonts";
const urls = new Set<string>();

/** Publica las fuentes subidas como @font-face, por encima de las sustitutas. */
export function installFontFaces(faces: ShelfFace[], aliases: Array<{ face: ShelfFace; family: string }>) {
  for (const url of urls) URL.revokeObjectURL(url);
  urls.clear();
  const rules: string[] = [];
  const declared = [...faces.map((face) => ({ face, family: face.family })), ...aliases];
  for (const { face, family } of declared) {
    const copy = new Uint8Array(face.bytes);
    const url = URL.createObjectURL(new Blob([copy], { type: "font/ttf" }));
    urls.add(url);
    const safe = family.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    rules.push(
      `@font-face{font-family:"${safe}";src:url("${url}") format("truetype");font-weight:${face.weight};font-style:${face.style};font-display:swap;}`
    );
  }
  let style = document.getElementById(STYLE_ID) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = STYLE_ID;
  }
  style.textContent = rules.join("\n");
  document.head.appendChild(style);
}

export async function registerShelfFonts(faces: ShelfFace[], aliases: Array<{ face: ShelfFace; family: string }>) {
  const { registerFont, clearFonts, clearDiscoveredFonts } = await import("jsafa-pdf");
  clearFonts();
  clearDiscoveredFonts();
  const jobs = [...faces.map((face) => ({ face, family: face.family })), ...aliases];
  await Promise.all(
    jobs.map(({ face, family }) =>
      registerFont({
        family,
        src: new Uint8Array(face.bytes),
        weight: face.weight,
        style: face.style,
      }),
    ),
  );
}

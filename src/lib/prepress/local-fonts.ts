import { fontKey } from "./font-face.ts";

/** Lo que da el navegador por cada cara instalada (Local Font Access: Chrome y Edge, con permiso de la persona). */
type LocalFontData = {
  family: string;
  fullName: string;
  postscriptName: string;
  style: string;
  blob(): Promise<Blob>;
};

type WithLocalFonts = { queryLocalFonts?: () => Promise<LocalFontData[]> };

/** Si este navegador puede leer las fuentes instaladas en la computadora (Chrome y Edge sí; Firefox y Safari no). */
export function canUseLocalFonts(): boolean {
  return typeof window !== "undefined" && typeof (window as unknown as WithLocalFonts).queryLocalFonts === "function";
}

export type LocalFonts = { files: File[]; found: string[]; missing: string[] };

/**
 * Las fuentes de esas familias que hay instaladas en la computadora, como archivos (las mismas que usa Word). Se piden
 * al navegador, que muestra su propio permiso; quedan solo en este navegador, no se suben a ningún lado. Tiene que
 * llamarse desde un clic. Si el navegador no puede o no da permiso, lanza un error con el motivo en palabras.
 */
export async function localFontFiles(families: string[]): Promise<LocalFonts> {
  const query = (window as unknown as WithLocalFonts).queryLocalFonts;
  if (typeof query !== "function") {
    throw new Error("Este navegador no puede leer las fuentes de la computadora: usá Chrome o Edge, o subí los archivos de las fuentes.");
  }
  let fonts: LocalFontData[];
  try {
    fonts = await query.call(window);
  } catch (err) {
    const name = err instanceof DOMException ? err.name : "";
    throw new Error(
      name === "NotAllowedError" || name === "SecurityError"
        ? "El navegador no dio permiso para leer las fuentes de la computadora. Si te lo vuelve a preguntar, elegí «Permitir»."
        : "No se pudieron leer las fuentes de la computadora.",
    );
  }
  const wanted = new Map(families.map((family) => [fontKey(family), family]));
  const files: File[] = [];
  const found = new Set<string>();
  // Varias caras de una misma colección (cambria.ttc: Cambria y Cambria Math) traen el mismo archivo: va una vez.
  const seen = new Set<string>();
  for (const font of fonts) {
    const family = wanted.get(fontKey(font.family));
    if (!family) continue;
    const blob = await font.blob();
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const id = `${bytes.length}:${bytes.subarray(0, 64).join(",")}:${bytes.subarray(-64).join(",")}`;
    found.add(family);
    if (seen.has(id)) continue;
    seen.add(id);
    const name = (font.postscriptName || font.fullName || font.family).replace(/[^\w.-]+/g, "-");
    files.push(new File([bytes], `${name}.ttf`, { type: "font/ttf" }));
  }
  return { files, found: [...found], missing: families.filter((family) => !found.has(family)) };
}

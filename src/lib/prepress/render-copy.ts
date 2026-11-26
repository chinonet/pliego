/**
 * Ajustes del .docx que solo sirven para componer (la copia que ve el visor, nunca el Word que se descarga) y datos del
 * diseño de página que el visor no aplica y que se completan después de dibujar.
 */
export type SectionLayout = {
  /** Cómo empieza la sección (w:type del sectPr). Ausente = página siguiente. */
  type?: "nextPage" | "oddPage" | "evenPage" | "continuous" | "nextColumn";
  /** Margen de encuadernación (w:gutter), en puntos. */
  gutterPt?: number;
  /** Márgenes izquierdo y derecho del Word, en puntos. */
  leftPt?: number;
  rightPt?: number;
};

export type RenderInfo = {
  /** Márgenes simétricos (settings.xml w:mirrorMargins): en las páginas pares se invierten izquierda y derecha. */
  mirrorMargins: boolean;
  /** Encuadernación arriba (w:gutterAtTop). */
  gutterAtTop: boolean;
  /** División automática de palabras activada en el Word (w:autoHyphenation). */
  autoHyphenation: boolean;
  /** Numeración de notas al pie (w:footnotePr en settings o en el sectPr final). */
  footnoteRestart: "continuous" | "eachSect" | "eachPage";
  footnoteStart: number;
  /** Una entrada por sección, en el mismo orden que SectionNumbering de fields.ts. */
  sections: SectionLayout[];
};

export const DEFAULT_RENDER_INFO: RenderInfo = {
  mirrorMargins: false,
  gutterAtTop: false,
  autoHyphenation: false,
  footnoteRestart: "continuous",
  footnoteStart: 1,
  sections: [],
};

/** Recibe la copia ya marcada por markPageFields y la deja lista para el visor. */
export async function prepareRenderCopy(
  bytes: Uint8Array,
): Promise<{ bytes: Uint8Array; info: RenderInfo; warnings: string[] }> {
  return { bytes, info: { ...DEFAULT_RENDER_INFO }, warnings: [] };
}

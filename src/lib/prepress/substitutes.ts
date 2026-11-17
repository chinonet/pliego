/** Familias métricas de uso libre. Solo se aplican si el usuario lo pide. */
export const SUBSTITUTE_FACES: Array<{
  family: string;
  file: string;
  weight: number;
  style: "normal" | "italic";
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
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Regular.ttf", weight: 400, style: "normal" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Bold.ttf", weight: 700, style: "normal" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-Italic.ttf", weight: 400, style: "italic" },
  { family: "Times New Roman", file: "/fonts/LiberationSerif-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Times", file: "/fonts/LiberationSerif-Regular.ttf", weight: 400, style: "normal" },
  { family: "Courier New", file: "/fonts/LiberationMono-Regular.ttf", weight: 400, style: "normal" },
  { family: "Courier New", file: "/fonts/LiberationMono-Bold.ttf", weight: 700, style: "normal" },
  { family: "Courier New", file: "/fonts/LiberationMono-Italic.ttf", weight: 400, style: "italic" },
  { family: "Courier New", file: "/fonts/LiberationMono-BoldItalic.ttf", weight: 700, style: "italic" },
  { family: "Tahoma", file: "/fonts/Tahoma.ttf", weight: 400, style: "normal" },
  { family: "Tahoma", file: "/fonts/Tahoma-Bold.ttf", weight: 700, style: "normal" },
];

const STYLE_ID = "pliego-substitutes";

export function installSubstitutes(enabled: boolean) {
  const existing = document.getElementById(STYLE_ID);
  if (!enabled) {
    existing?.remove();
    return;
  }
  const style = existing ?? document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = SUBSTITUTE_FACES.map(
    (face) =>
      `@font-face{font-family:"${face.family}";src:url("${face.file}") format("truetype");font-weight:${face.weight};font-style:${face.style};font-display:swap;}`,
  ).join("\n");
  document.head.appendChild(style);
}

export async function loadSubstituteFaces() {
  if (!("fonts" in document)) return;
  await Promise.all(
    SUBSTITUTE_FACES.map((face) =>
      document.fonts.load(`${face.style} ${face.weight} 16px "${face.family}"`).catch(() => undefined),
    ),
  );
  await document.fonts.ready;
}

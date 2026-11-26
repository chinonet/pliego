/**
 * Arreglos a las reglas de estilo que escribe el visor (docx-preview).
 *
 * El estilo predeterminado del Word (por ejemplo «Normal») sale como «.docx p, p.docx_normal span { … }»: la letra y el
 * tamaño del estilo quedan en el párrafo y no en el texto, y entonces la configuración general del documento
 * («.docx span { … }», docDefaults) le gana en cada tramo. En Word es al revés: el estilo manda sobre la configuración
 * general. Se corrige a «.docx p span, p.docx_normal span».
 */
export function fixDefaultStyleSelectors(css: string, className = "docx"): string {
  const escaped = className.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`\\.${escaped} ([a-z0-9]+), \\1\\.([\\w-]+) ([a-z0-9]+)(\\s*\\{)`, "gi");
  return css.replace(pattern, (_all, target: string, name: string, sub: string, open: string) =>
    `.${className} ${target} ${sub}, ${target}.${name} ${sub}${open}`,
  );
}

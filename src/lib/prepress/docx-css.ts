/**
 * Arreglos a las reglas de estilo que escribe el visor (docx-preview).
 *
 * El estilo predeterminado del Word (por ejemplo «Normal») sale como «.docx p, p.docx_normal span { … }»: la letra y el
 * tamaño del estilo quedan en el párrafo y no en el texto, y entonces la configuración general del documento
 * («.docx span { … }», docDefaults) le gana en cada tramo. En Word es al revés: el estilo manda sobre la configuración
 * general, y los estilos de carácter («Código», «Hipervínculo»…) mandan sobre el estilo del párrafo.
 *
 * Se corrige a «.docx :where(p:not([class*="docx_"]), p.docx_normal) span»: vale para los párrafos sin estilo (que en
 * Word son «Normal») y los que dicen «Normal», no para los que tienen otro estilo (el visor ya les copia lo heredado de
 * «Normal» en su propia regla). Pesa lo mismo que «.docx span» y que «span.docx_codigo»; como el visor escribe primero
 * la configuración general y después los estilos en el orden del Word, gana sobre la primera y pierde con los estilos
 * de carácter, igual que en Word.
 */
export function fixDefaultStyleSelectors(css: string, className = "docx"): string {
  const escaped = className.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`\\.${escaped} ([a-z0-9]+), \\1\\.([\\w-]+) ([a-z0-9]+)(\\s*\\{)`, "gi");
  return css.replace(pattern, (_all, target: string, name: string, sub: string, open: string) =>
    `.${className} :where(${target}:not([class*="${className}_"]), ${target}.${name}) ${sub}${open}`,
  );
}

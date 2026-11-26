import assert from "node:assert/strict";
import { fixDefaultStyleSelectors } from "./docx-css.ts";

assert.equal(
  fixDefaultStyleSelectors(".docx p, p.docx_normal span {font-family: Cambria; font-size: 12pt}"),
  '.docx :where(p:not([class*="docx_"]), p.docx_normal) span {font-family: Cambria; font-size: 12pt}',
  "la letra del estilo predeterminado va al texto",
);
assert.equal(
  fixDefaultStyleSelectors(".docx p, p.docx_normal {margin-bottom: 10pt}"),
  ".docx p, p.docx_normal {margin-bottom: 10pt}",
  "lo que es del párrafo queda en el párrafo",
);
assert.equal(
  fixDefaultStyleSelectors(".docx span {font-family: Calibri}"),
  ".docx span {font-family: Calibri}",
  "la configuración general no cambia",
);
assert.equal(
  fixDefaultStyleSelectors(".docx table, table.docx_tablanormal td {padding: 0}"),
  '.docx :where(table:not([class*="docx_"]), table.docx_tablanormal) td {padding: 0}',
);
assert.equal(
  fixDefaultStyleSelectors(".pv p, p.pv_normal span {color: #000}", "pv"),
  '.pv :where(p:not([class*="pv_"]), p.pv_normal) span {color: #000}',
);

// Peso de cada regla (ids, clases y atributos, etiquetas), contado a mano para estos casos: lo de adentro de :where no
// cuenta. «Normal» tiene que pesar igual que la configuración general y que un estilo de carácter (así decide el orden
// en que el visor los escribe: general, Normal, Código), y no más que el estilo de carácter.
function weight(selector: string): [number, number, number] {
  const outside = selector.replace(/:where\((?:[^()]|\([^()]*\))*\)/g, "");
  const classes = (outside.match(/\.[\w-]+|\[[^\]]*\]|:(?!where)[\w-]+/g) ?? []).length;
  const tags = (outside.match(/(?:^|[\s>+~])([a-z][a-z0-9]*)/gi) ?? []).length;
  return [0, classes, tags];
}
const normal = fixDefaultStyleSelectors(".docx p, p.docx_normal span {font-family: Garamond}").split("{")[0]!.trim();
const general = ".docx span";
const codigo = "span.docx_codigo";
assert.deepEqual(weight(normal), weight(codigo), "un estilo de carácter dentro de un párrafo Normal no pierde por peso");
assert.deepEqual(weight(normal), weight(general), "Normal le gana a la configuración general por venir después");
assert.deepEqual(weight("p.docx_titulo span"), [0, 1, 2], "el peso se cuenta bien");

console.log("docx-css ok");

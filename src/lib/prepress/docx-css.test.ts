import assert from "node:assert/strict";
import { fixDefaultStyleSelectors } from "./docx-css.ts";

assert.equal(
  fixDefaultStyleSelectors(".docx p, p.docx_normal span {font-family: Cambria; font-size: 12pt}"),
  ".docx p span, p.docx_normal span {font-family: Cambria; font-size: 12pt}",
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
  ".docx table td, table.docx_tablanormal td {padding: 0}",
);

console.log("docx-css ok");

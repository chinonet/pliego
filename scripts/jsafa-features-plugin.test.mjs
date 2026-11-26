import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import { PDF_FONT_FEATURES, patchJsafa } from "./jsafa-features-plugin.mjs";

test("pone las features en las llamadas de jsafa que incrustan fuentes", () => {
  const sample = "a = await t.embedFont(a.bytes, { subset: r }); b = t.embedFont(o.bytes, { subset: !1 }); c = t.embedFont(o.name);";
  const { code, count } = patchJsafa(sample);
  assert.equal(count, 2);
  assert.ok(code.includes(`t.embedFont(a.bytes, { subset: r, features: ${PDF_FONT_FEATURES} })`));
  assert.ok(code.includes(`t.embedFont(o.bytes, { subset: !1, features: ${PDF_FONT_FEATURES} })`));
  assert.ok(code.includes("t.embedFont(o.name)"), "las fuentes estándar del PDF quedan igual");
});

test("el jsafa-pdf instalado tiene las cuatro llamadas", async () => {
  const path = new URL("../node_modules/jsafa-pdf/dist/jsafa-pdf.js", import.meta.url);
  const source = await readFile(path, "utf8");
  const { code, count } = patchJsafa(source);
  assert.equal(count, 4);
  assert.equal(patchJsafa(code).count, 0, "no se aplica dos veces");
});

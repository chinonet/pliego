import assert from "node:assert/strict";
import JSZip from "jszip";
import { markPageFields, PAGE_MARK, PAGES_MARK, sectionMark } from "./fields.ts";
import { formatPageNumber } from "./paginate.ts";
import { neutralToGray } from "./export-pdf.ts";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p><w:r><w:t>Prólogo</w:t></w:r></w:p>
<w:p><w:pPr><w:sectPr><w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr><w:r><w:t>Fin del prólogo</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Capítulo 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgNumType w:start="1"/></w:sectPr>
</w:body></w:document>`;
const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${W}><w:p>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>2</w:t></w:r><w:r><w:t>0</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:t xml:space="preserve"> de </w:t></w:r>
<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>`;

const zip = new JSZip();
zip.file("word/document.xml", docXml);
zip.file("word/footer1.xml", footerXml);
zip.file("word/media/image1.png", new Uint8Array([1, 2, 3, 4]));
const input = await zip.generateAsync({ type: "uint8array" });

const marked = await markPageFields(input);
const out = await JSZip.loadAsync(marked.bytes);
const footer = await out.file("word/footer1.xml")!.async("string");
const body = await out.file("word/document.xml")!.async("string");

assert.equal(marked.fields, 3, "dos PAGE y un NUMPAGES");
assert.equal(footer.split(PAGE_MARK).length - 1, 2, "cada PAGE queda como una marca");
assert.ok(footer.includes(PAGES_MARK), "NUMPAGES queda como marca");
assert.doesNotMatch(footer, /<w:t[^>]*>2<\/w:t>|<w:t[^>]*>0<\/w:t>|<w:t[^>]*>9<\/w:t>/, "sin los números guardados por Word");
assert.doesNotMatch(footer, /fldSimple/, "el visor no dibuja fldSimple: queda como texto");
assert.match(footer, /<w:b\/>/, "se conserva el formato del número");
assert.deepEqual(marked.sections, [{ start: undefined, format: "lowerRoman" }, { start: 1, format: undefined }]);
assert.ok(body.indexOf(sectionMark(0)) < body.indexOf("Prólogo"), "marca al comienzo de la primera sección");
assert.ok(body.indexOf(sectionMark(1)) < body.indexOf("Capítulo 1"), "marca dentro de la tabla que abre la segunda sección");
assert.deepEqual(
  await out.file("word/media/image1.png")!.async("uint8array"),
  new Uint8Array([1, 2, 3, 4]),
  "las imágenes no se tocan",
);

assert.equal(formatPageNumber(4, "lowerRoman"), "iv");
assert.equal(formatPageNumber(14, "upperRoman"), "XIV");
assert.equal(formatPageNumber(28, "lowerLetter"), "bb");
assert.equal(formatPageNumber(7), "7");

assert.equal(neutralToGray("BT\n0 0 0 rg\n/F1 12 Tf"), "BT\n0 g\n/F1 12 Tf");
assert.equal(neutralToGray("0.2 0.2 0.2 RG 1 w"), "0.2 G 1 w");
assert.equal(neutralToGray("1 0 0 rg"), "1 0 0 rg", "un color de verdad no se toca");
assert.equal(neutralToGray("10 0 0 10 0 0 cm 0 0 0 rg"), "10 0 0 10 0 0 cm 0 g", "solo los operandos de rg");

console.log("fields ok");

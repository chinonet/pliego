import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { cleanDocx, isFontName } from "./clean-docx.ts";

const original = new Uint8Array(await readFile(new URL("../../../public/muestra-premedia.docx", import.meta.url)));
const source = await JSZip.loadAsync(original);
const sourceImage = await source.file(/word\/media\//)[0]?.async("uint8array");
assert.ok(sourceImage && sourceImage.byteLength > 1000, "la muestra trae imagen");

const cleaned = await cleanDocx(original, {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: true,
  removeHighlights: false,
});

const zip = await JSZip.loadAsync(cleaned.bytes);
assert.equal(zip.file("word/comments.xml"), null);
assert.equal(zip.file(/comments/i).length, 0);

const documentXml = await zip.file("word/document.xml")!.async("string");
assert.match(documentXml, /Tirada de 2\.000 ejemplares/);
assert.doesNotMatch(documentXml, /Borrador interno/);
assert.doesNotMatch(documentXml, /<w:ins[\s>]/);
assert.doesNotMatch(documentXml, /<w:del[\s>]/);
assert.doesNotMatch(documentXml, /commentReference/);
assert.doesNotMatch(documentXml, /commentRangeStart/);
assert.match(documentXml, /<w:footnoteReference\b/);
assert.match(documentXml, /<w:endnoteReference\b/);
assert.match(documentXml, /w:val="yellow"/);

const footnotes = await zip.file("word/footnotes.xml")!.async("string");
assert.match(footnotes, /papel estucado de 150 g/);

const image = await zip.file(/word\/media\//)[0]!.async("uint8array");
assert.equal(image.byteLength, sourceImage.byteLength);
assert.equal(Buffer.from(image).equals(Buffer.from(sourceImage)), true);

assert.equal(cleaned.report.commentsRemoved, 1);
assert.equal(cleaned.report.footnotesKept, 1);
assert.equal(cleaned.report.endnotesKept, 1);
assert.ok(cleaned.report.insertionsAccepted >= 1);
assert.ok(cleaned.report.deletionsRemoved >= 1);
assert.equal(cleaned.report.images.length, 1);
assert.equal(isFontName("Calibri"), true);
assert.equal(isFontName("MS Gothic"), true);
assert.equal(isFontName("inherit"), false);
assert.equal(isFontName("es-AR"), false);

const noMarks = await cleanDocx(original, {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: false,
  removeHighlights: true,
});
const bare = await JSZip.loadAsync(noMarks.bytes);
const bareXml = await bare.file("word/document.xml")!.async("string");
assert.doesNotMatch(bareXml, /<w:endnoteReference\b/);
assert.doesNotMatch(bareXml, /w:val="yellow"/);
assert.match(bareXml, /<w:footnoteReference\b/);

console.log("clean-docx ok", cleaned.report);

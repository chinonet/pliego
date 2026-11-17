import { deflateSync, crc32 } from "node:zlib";
import { writeFile } from "node:fs/promises";
import {
  AlignmentType,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  EndnoteReferenceRun,
  Footer,
  FootnoteReferenceRun,
  Header,
  HeadingLevel,
  ImageRun,
  InsertedTextRun,
  Packer,
  PageNumber,
  Paragraph,
  TextRun,
} from "docx";

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const name = Buffer.from(type);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(Number(crc32(Buffer.concat([name, data]))) >>> 0);
  return Buffer.concat([len, name, data, crc]);
}

/** Trama fina + barras + círculo. Si el PDF la suaviza, hubo recompressión. */
function makePng(w, h) {
  const raw = Buffer.alloc((w * 4 + 1) * h);
  const cx = Math.round(w * 0.72);
  const cy = Math.round(h * 0.58);
  const radius = Math.round(h * 0.28);
  for (let y = 0; y < h; y++) {
    const row = y * (w * 4 + 1);
    raw[row] = 0;
    for (let x = 0; x < w; x++) {
      const i = row + 1 + x * 4;
      let r = 244;
      let g = 239;
      let b = 230;
      const bar = Math.floor((x / w) * 8);
      if (y < 36) {
        const bars = [
          [20, 20, 20],
          [196, 40, 28],
          [20, 90, 160],
          [20, 130, 70],
          [230, 190, 40],
          [120, 40, 140],
          [240, 240, 240],
          [80, 80, 80],
        ];
        [r, g, b] = bars[bar];
      } else if (x < 140 && y > 56 && y < h - 24) {
        const on = ((x + y) & 1) === 0;
        r = g = b = on ? 16 : 235;
      }
      const dx = x - cx;
      const dy = y - cy;
      if (dx * dx + dy * dy < radius * radius) {
        r = 30;
        g = 78;
        b = 140;
      }
      if (dx * dx + dy * dy < (radius - 10) * (radius - 10) && dx * dx + dy * dy > (radius - 14) * (radius - 14)) {
        r = g = b = 250;
      }
      raw[i] = r;
      raw[i + 1] = g;
      raw[i + 2] = b;
      raw[i + 3] = 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  return Buffer.concat([
    sig,
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const png = makePng(640, 360);

const doc = new Document({
  title: "Muestra premedia",
  creator: "Pliego",
  description: "Documento de prueba con marcas, comentario, nota al pie e imagen.",
  features: { trackRevisions: true },
  comments: {
    children: [
      {
        id: 0,
        author: "Mesa de color",
        initials: "MC",
        date: new Date("2026-04-02T11:00:00Z"),
        children: [
          new Paragraph("Comentario de revisión: no debe salir en el PDF de imprenta."),
        ],
      },
    ],
  },
  footnotes: {
    1: {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              font: "Calibri",
              text: "Nota al pie conservada: papel estucado de 150 g, cara y dorso.",
            }),
          ],
        }),
      ],
    },
  },
  endnotes: {
    1: {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              font: "Calibri",
              text: "Nota al final conservada: perfil de salida ISO coated v2, sin convertir en este paso.",
            }),
          ],
        }),
      ],
    },
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 720, bottom: 720, left: 900, right: 900 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  font: "Calibri",
                  italics: true,
                  size: 18,
                  text: "Pliego  ·  cuaderno de premedia",
                  color: "5C564E",
                }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({ font: "Calibri", size: 18, text: "Página " }),
                new TextRun({ font: "Calibri", size: 18, children: [PageNumber.CURRENT] }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [
            new TextRun({
              font: "Cambria",
              text: "Ficha de impresión",
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({
              font: "Calibri",
              size: 24,
              text: "Cubierta en cuatricromía. Texto firme que no debe moverse. ",
            }),
            new InsertedTextRun({
              id: 1,
              author: "Edición",
              date: "2026-04-02T10:00:00Z",
              font: "Calibri",
              size: 24,
              text: "Tirada de 2.000 ejemplares.",
            }),
            new TextRun({ font: "Calibri", size: 24, text: " " }),
            new DeletedTextRun({
              id: 2,
              author: "Edición",
              date: "2026-04-02T10:04:00Z",
              font: "Calibri",
              size: 24,
              text: "Borrador interno que no va a máquina.",
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({ font: "Calibri", size: 24, text: "El pie de foto lleva " }),
            new CommentRangeStart(0),
            new TextRun({
              font: "Calibri",
              size: 24,
              text: "una frase comentada",
              highlight: "yellow",
            }),
            new CommentRangeEnd(0),
            new CommentReference(0),
            new TextRun({ font: "Calibri", size: 24, text: " y una nota al pie" }),
            new FootnoteReferenceRun(1),
            new TextRun({ font: "Calibri", size: 24, text: " más una nota al final" }),
            new EndnoteReferenceRun(1),
            new TextRun({ font: "Calibri", size: 24, text: "." }),
          ],
        }),
        new Paragraph({
          spacing: { before: 200, after: 80 },
          children: [
            new TextRun({
              font: "Calibri",
              size: 22,
              text: "Imagen de control. La trama de 1 px y el aro tienen que salir nítidos: el conversor no recomprime.",
            }),
          ],
        }),
        new Paragraph({
          children: [
            new ImageRun({
              type: "png",
              data: png,
              transformation: { width: 460, height: 259 },
              altText: { name: "control", title: "Trama de control", description: "Barras, damero y aro" },
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
await writeFile(new URL("../public/muestra-premedia.docx", import.meta.url), buffer);
console.log("wrote muestra-premedia.docx", buffer.byteLength);

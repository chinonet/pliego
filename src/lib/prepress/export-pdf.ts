/**
 * Cada sección ya paginada por el visor se vuelca a una página PDF del mismo
 * tamaño. El texto sigue siendo texto (fuentes incrustadas) y un JPEG del
 * Word entra como DCT, sin reencodear.
 */
export async function exportPagesToPdf(
  sections: HTMLElement[],
  opts: { title: string; subset: boolean; cropMarks?: boolean },
): Promise<Uint8Array> {
  const { toPdf } = await import("jsafa-pdf");
  const { PDFDocument, rgb } = await import("pdf-lib");

  if (sections.length === 0) throw new Error("El documento no tiene páginas para exportar.");

  const merged = await PDFDocument.create();
  merged.setTitle(opts.title);
  merged.setCreator("Pliego");
  merged.setProducer("Pliego premedia");

  for (const section of sections) {
    const bytes = await toPdf(section, {
      subset: opts.subset,
      title: opts.title,
      autoFonts: true,
      waitForAssets: true,
    });
    const part = await PDFDocument.load(bytes);
    const copied = await merged.copyPages(part, part.getPageIndices());
    for (const page of copied) {
      if (opts.cropMarks) drawCropMarks(page, rgb);
      merged.addPage(page);
    }
  }

  return merged.save();
}

/** Marcas de corte en vectores, fuera del recorte. No tocan las imágenes. */
function drawCropMarks(
  page: import("pdf-lib").PDFPage,
  rgb: typeof import("pdf-lib").rgb,
) {
  const width = page.getWidth();
  const height = page.getHeight();
  const gap = 6;
  const len = 14;
  const pad = gap + len;
  page.setMediaBox(-pad, -pad, width + pad * 2, height + pad * 2);
  const ink = rgb(0, 0, 0);
  const lines: Array<[[number, number], [number, number]]> = [
    [[-pad, 0], [-gap, 0]],
    [[0, -pad], [0, -gap]],
    [[width + gap, 0], [width + pad, 0]],
    [[width, -pad], [width, -gap]],
    [[-pad, height], [-gap, height]],
    [[0, height + gap], [0, height + pad]],
    [[width + gap, height], [width + pad, height]],
    [[width, height + gap], [width, height + pad]],
  ];
  for (const [start, end] of lines) {
    page.drawLine({
      start: { x: start[0], y: start[1] },
      end: { x: end[0], y: end[1] },
      thickness: 0.4,
      color: ink,
    });
  }
}

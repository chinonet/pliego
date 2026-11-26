import { useEffect, useRef, useState, type ReactNode } from "react";
import { flushSync } from "react-dom";
import {
  AlertTriangle,
  FileDown,
  FileText,
  ImageIcon,
  Pilcrow,
  Scissors,
  Type,
  Upload,
} from "lucide-react";
import { cleanDocx, type CleanOptions, type CleanReport } from "@/lib/prepress/clean-docx";
import { fontKey } from "@/lib/prepress/font-face";
import { forgetFace, installFontFaces, loadShelf, readFontFile, rememberFace, registerShelfFonts, type ShelfFace } from "@/lib/prepress/font-shelf";
import { installSubstitutes, loadSubstituteFaces, SUBSTITUTE_FACES } from "@/lib/prepress/substitutes";
import { exportPagesToPdf } from "@/lib/prepress/export-pdf";
import { markPageFields } from "@/lib/prepress/fields";
import { numberPages, repaginate, takeSectionMarks } from "@/lib/prepress/paginate";

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function saveBytes(bytes: Uint8Array, filename: string, type: string) {
  const copy = new Uint8Array(bytes);
  const blob = new Blob([copy], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function stem(name: string): string {
  return name.replace(/\.docx$/i, "") || "documento";
}

function Toggle({
  pressed,
  label,
  hint,
  onClick,
}: {
  pressed: boolean;
  label: string;
  hint: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      aria-pressed={pressed}
      onClick={onClick}
      className="flex w-full items-start gap-3 rounded-lg border border-line bg-surface px-3 py-3 text-left"
    >
      <span
        className={
          "mt-0.5 grid size-5 shrink-0 place-items-center rounded border " +
          (pressed ? "border-accent bg-accent text-bg" : "border-line bg-surface-2 text-transparent")
        }
        aria-hidden="true"
      >
        <svg viewBox="0 0 16 16" className="size-3.5">
          <path d="M3 8.2 6.2 11.5 13 4.5" fill="none" stroke="currentColor" strokeWidth="2" />
        </svg>
      </span>
      <span>
        <span className="block text-sm font-semibold text-fg">{label}</span>
        <span className="mt-0.5 block text-sm leading-snug text-muted">{hint}</span>
      </span>
    </button>
  );
}

export function Studio() {
  const inputRef = useRef<HTMLInputElement>(null);
  const frameRef = useRef<HTMLDivElement>(null);
  const docRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const fontRef = useRef<HTMLInputElement>(null);
  const shelfRef = useRef<ShelfFace[]>([]);
  const reportRef = useRef<CleanReport | null>(null);
  const cleanedRef = useRef<Uint8Array | null>(null);
  const sourceRef = useRef<Uint8Array | null>(null);
  const nameRef = useRef("documento.docx");

  const [name, setName] = useState("");
  const [phase, setPhase] = useState<"idle" | "working" | "ready" | "error">("idle");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [report, setReport] = useState<CleanReport | null>(null);
  const [pageCount, setPageCount] = useState(0);
  const [fit, setFit] = useState(true);
  const [scale, setScale] = useState(1);
  const [sheet, setSheet] = useState({ width: 0, height: 0 });
  const [options, setOptions] = useState<CleanOptions>({
    acceptRevisions: true,
    removeComments: true,
    keepEndnotes: true,
    removeHighlights: false,
  });
  const [fullFont, setFullFont] = useState(true);
  const [substitutes, setSubstitutes] = useState(true);
  const substitutesRef = useRef(true);
  const [cropMarks, setCropMarks] = useState(false);
  const [grayscale, setGrayscale] = useState(false);
  const [press, setPress] = useState<{ pdfx: boolean; problems: string[] } | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [shelf, setShelf] = useState<ShelfFace[]>([]);
  const [imageChecks, setImageChecks] = useState<Array<{ label: string; ppi: number; pixels: string }>>([]);
  const [pageMm, setPageMm] = useState("");
  const [tooTall, setTooTall] = useState<number[]>([]);

  function measure() {
    const doc = docRef.current;
    const frame = frameRef.current;
    const wrapper = doc?.querySelector(".docx-wrapper") as HTMLElement | null;
    const page = doc?.querySelector("section.docx") as HTMLElement | null;
    if (!doc || !frame || !wrapper || !page) return;
    const width = page.offsetWidth;
    const height = wrapper.offsetHeight;
    const avail = Math.max(frame.clientWidth - 48, 160);
    const next = fit ? Math.min(1, avail / width) : 1;
    setSheet({ width, height });
    setScale(Number.isFinite(next) && next > 0 ? next : 1);
  }

  function fontAliases(faces: ShelfFace[], current: CleanReport | null) {
    if (!current) return [];
    const needed = current.fonts;
    const aliases: Array<{ face: ShelfFace; family: string }> = [];
    for (const face of faces) {
      const stem = fontKey(face.fileName.replace(/\.[^.]+$/, ""));
      const byName = needed.filter((font) => fontKey(face.family) === fontKey(font.name));
      const byFile = needed
        .filter((font) => {
          const key = fontKey(font.name);
          return stem === key || stem.startsWith(key);
        })
        .sort((a, b) => fontKey(b.name).length - fontKey(a.name).length);
      const targets = byName.length ? byName : byFile.slice(0, 1);
      for (const font of targets) aliases.push({ face, family: font.name });
    }
    return aliases;
  }

  useEffect(() => {
    void loadShelf().then((faces) => {
      shelfRef.current = faces;
      setShelf(faces);
    });
  }, []);

  useEffect(() => {
    const frame = frameRef.current;
    if (!frame) return;
    measure();
    const observer = new ResizeObserver(() => measure());
    observer.observe(frame);
    return () => observer.disconnect();
  }, [fit, pageCount, phase]);

  async function compose(bytes: Uint8Array, clean: CleanOptions) {
    setStatus("Componiendo las páginas");
    installSubstitutes(substitutesRef.current);
    const aliases = fontAliases(shelfRef.current, reportRef.current);
    installFontFaces(shelfRef.current, aliases);
    if (substitutesRef.current) await loadSubstituteFaces();
    if ("fonts" in document) {
      const families = new Set<string>(aliases.map((item) => item.family));
      for (const face of shelfRef.current) families.add(face.family);
      await Promise.all(
        [...families].flatMap((family) =>
          shelfRef.current
            .filter((face) => face.family === family || aliases.some((item) => item.family === family && item.face === face))
            .map((face) => document.fonts.load(`${face.style} ${face.weight} 16px "${family}"`).catch(() => undefined)),
        ),
      );
      await document.fonts.ready;
    }
    const { renderAsync } = await import("docx-preview");
    const body = docRef.current;
    const style = styleRef.current;
    if (!body || !style) throw new Error("No está el visor.");
    body.replaceChildren();
    style.replaceChildren();
    // Copia solo para componer: los números de página se escriben después, página por página.
    const marked = await markPageFields(bytes);
    await renderAsync(marked.bytes, body, style, {
      className: "docx",
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: clean.keepEndnotes,
      renderComments: false,
      renderChanges: false,
      useBase64URL: false,
      experimental: true,
    });
    const covered = new Set<string>();
    if (substitutesRef.current) {
      for (const face of SUBSTITUTE_FACES) covered.add(face.family.toLowerCase());
    }
    for (const face of shelfRef.current) covered.add(face.family.toLowerCase());
    for (const alias of aliases) covered.add(alias.family.toLowerCase());
    // Solo se quitan las fuentes incrustadas (@font-face) que la app reemplaza. Los demás bloques son los estilos,
    // el tema y la numeración del Word: si se borran, el documento pierde su letra, sus tamaños y sus sangrías.
    for (const node of [...style.querySelectorAll("style")]) {
      const css = (node.textContent ?? "").trim();
      if (!/^@font-face/i.test(css)) continue;
      const family = /font-family:\s*([^;}]+)/i.exec(css)?.[1]?.trim().replace(/^['"]+|['"]+$/g, "").toLowerCase();
      if (family && covered.has(family)) node.remove();
    }
    for (const id of ["pliego-substitutes", "pliego-uploaded-fonts"]) {
      const tag = document.getElementById(id);
      if (tag) document.head.appendChild(tag);
    }
    if (substitutesRef.current) await loadSubstituteFaces();
    if ("fonts" in document && covered.size > 0) {
      await Promise.all(
        [...covered].map((family) => document.fonts.load(`16px "${family}"`).catch(() => undefined)),
      );
      await document.fonts.ready;
    }
    const images = [...body.querySelectorAll("section.docx img")] as HTMLImageElement[];
    await Promise.all(
      images.map(
        (img) =>
          img.complete && img.naturalWidth > 0
            ? Promise.resolve()
            : new Promise<void>((resolve) => {
                img.addEventListener("load", () => resolve(), { once: true });
                img.addEventListener("error", () => resolve(), { once: true });
              }),
      ),
    );
    const starts = takeSectionMarks(body);
    const flow = repaginate(body);
    numberPages(body, marked.sections, starts);
    setTooTall(flow.tooTall);
    const pages = body.querySelectorAll("section.docx").length;
    setPageCount(pages);
    const first = body.querySelector("section.docx") as HTMLElement | null;
    if (first) {
      const width = Math.round((first.offsetWidth / 96) * 25.4);
      const height = Math.round((first.offsetHeight / 96) * 25.4);
      setPageMm(`${width} × ${height} mm`);
    }
    setImageChecks(
      [...body.querySelectorAll("section.docx img")].map((node, index) => {
        const img = node as HTMLImageElement;
        const inches = img.clientWidth / 96;
        const ppi = inches > 0 ? Math.round(img.naturalWidth / inches) : 0;
        return {
          label: img.alt || `Imagen ${index + 1}`,
          ppi,
          pixels: `${img.naturalWidth}×${img.naturalHeight}`,
        };
      }),
    );
    requestAnimationFrame(measure);
  }

  async function process(input: Uint8Array, filename: string, clean: CleanOptions) {
    setPhase("working");
    setError("");
    setPress(null);
    setStatus("Leyendo el Word");
    nameRef.current = filename;
    setName(filename);
    sourceRef.current = input;
    try {
      if (input.byteLength < 4 || input[0] !== 0x50 || input[1] !== 0x4b) {
        throw new Error("Eso no es un .docx. El .doc antiguo hay que guardarlo como .docx en Word.");
      }
      setStatus("Quitando marcas y comentarios");
      const result = await cleanDocx(input, clean);
      cleanedRef.current = result.bytes;
      reportRef.current = result.report;
      setReport(result.report);
      await compose(result.bytes, clean);
      setPhase("ready");
      setStatus("");
    } catch (err) {
      setPhase("error");
      setStatus("");
      setError(err instanceof Error ? err.message : "No se pudo preparar el documento.");
    }
  }

  async function toggleSubstitutes() {
    const next = !substitutesRef.current;
    substitutesRef.current = next;
    setSubstitutes(next);
    const bytes = cleanedRef.current;
    if (!bytes) return;
    setPhase("working");
    setStatus(next ? "Aplicando sustitutas métricas" : "Dejando solo las fuentes reales");
    try {
      await compose(bytes, options);
      setPhase("ready");
      setStatus("");
    } catch (err) {
      setPhase("error");
      setStatus("");
      setError(err instanceof Error ? err.message : "No se pudo recomponer.");
    }
  }

  function isFontFile(fileName: string) {
    return /\.(ttf|otf|ttc|otc)$/i.test(fileName);
  }

  async function addFonts(files: File[]) {
    const next = [...shelfRef.current];
    const problems: string[] = [];
    for (const file of files) {
      try {
        const faces = await readFontFile(file);
        for (const face of faces) {
          const slot = `${fontKey(face.family)}:${face.weight}:${face.style}`;
          const index = next.findIndex((item) => `${fontKey(item.family)}:${item.weight}:${item.style}` === slot);
          if (index >= 0) next.splice(index, 1);
          next.push(face);
          await rememberFace(face);
        }
      } catch (err) {
        problems.push(err instanceof Error ? `${file.name}: ${err.message}` : file.name);
      }
    }
    shelfRef.current = next;
    setShelf(next);
    if (problems.length) setError(problems.join(" "));
    else setError("");
    const bytes = cleanedRef.current;
    if (!bytes) return;
    setPhase("working");
    setStatus("Aplicando las fuentes subidas");
    try {
      await compose(bytes, options);
      setPhase("ready");
      setStatus("");
    } catch (err) {
      setPhase("error");
      setStatus("");
      setError(err instanceof Error ? err.message : "No se pudo recomponer con esas fuentes.");
    }
  }

  async function removeUploaded(id: string) {
    await forgetFace(id).catch(() => undefined);
    const next = shelfRef.current.filter((face) => face.id !== id);
    shelfRef.current = next;
    setShelf(next);
    const bytes = cleanedRef.current;
    if (!bytes) return;
    setPhase("working");
    try {
      await compose(bytes, options);
      setPhase("ready");
    } catch (err) {
      setPhase("error");
      setError(err instanceof Error ? err.message : "No se pudo recomponer.");
    }
  }

  async function takeFile(file: File) {
    if (isFontFile(file.name)) {
      await addFonts([file]);
      return;
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setPhase("error");
      setError("Acepto un .docx o las fuentes que le faltan, en TTF u OTF.");
      return;
    }
    const input = new Uint8Array(await file.arrayBuffer());
    await process(input, file.name, options);
  }

  async function openSample() {
    setStatus("Cargando la muestra editorial");
    setPhase("working");
    const response = await fetch("/muestra-premedia.docx");
    if (!response.ok) {
      setPhase("error");
      setError("No se pudo abrir la muestra.");
      return;
    }
    const input = new Uint8Array(await response.arrayBuffer());
    await process(input, "muestra-premedia.docx", options);
  }

  function patchOptions(patch: Partial<CleanOptions>) {
    const next = { ...options, ...patch };
    setOptions(next);
    const source = sourceRef.current;
    if (source) void process(source, nameRef.current, next);
  }

  async function downloadDocx() {
    const bytes = cleanedRef.current;
    if (!bytes) return;
    saveBytes(
      bytes,
      `${stem(nameRef.current)}-limpio.docx`,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
  }

  async function downloadPdf() {
    const doc = docRef.current;
    if (!doc || phase !== "ready") return;
    setStatus("Incrustando fuentes y pasando las imágenes tal cual");
    const previousFit = fit;
    flushSync(() => {
      setFit(false);
      setScale(1);
    });
    // Sin el fondo de «papel» ni la sombra de la pantalla: en el PDF la hoja es blanca, sin un rectángulo crema.
    doc.classList.add("pliego-exporting");
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    try {
      const sections = [...doc.querySelectorAll("section.docx")] as HTMLElement[];
      await registerShelfFonts(shelfRef.current, fontAliases(shelfRef.current, reportRef.current));
      let result: Awaited<ReturnType<typeof exportPagesToPdf>>;
      try {
        result = await exportPagesToPdf(sections, {
          title: stem(nameRef.current),
          subset: !fullFont,
          cropMarks,
          grayscale,
        });
      } catch (err) {
        if (!fullFont) throw err;
        result = await exportPagesToPdf(sections, { title: stem(nameRef.current), subset: true, cropMarks, grayscale });
        setFullFont(false);
        setError("Esa fuente no permite incrustarse entera. Se incrustó el subconjunto usado.");
      }
      setPress({ pdfx: result.pdfx, problems: result.problems });
      saveBytes(result.bytes, `${stem(nameRef.current)}.pdf`, "application/pdf");
      setStatus("");
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "No se pudo escribir el PDF.",
      );
      setStatus("");
    } finally {
      doc.classList.remove("pliego-exporting");
      flushSync(() => setFit(previousFit));
      requestAnimationFrame(measure);
    }
  }

  const imageBytes = report?.images.reduce((sum, image) => sum + image.bytes, 0) ?? 0;
  const busy = phase === "working";

  return (
    <main className="flex min-h-screen flex-col bg-bg pb-28 text-fg">
      <header className="no-print flex items-end justify-between gap-4 border-b border-line px-4 py-4 sm:px-6">
        <div>
          <p className="font-display text-4xl leading-none tracking-tight">Pliego</p>
          <p className="mt-1 text-sm text-muted">Word a PDF para premedia</p>
        </div>
        <p className="max-w-xs text-right text-sm text-muted">El archivo se prepara en este navegador. No se sube.</p>
      </header>

      <div className="grid flex-1 lg:grid-cols-[22rem_minmax(0,1fr)]">
        <aside className="no-print flex flex-col gap-4 border-b border-line p-4 sm:p-5 lg:border-r lg:border-b-0">
          <div
            onDragOver={(event) => {
              event.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDragOver(false);
              const files = [...event.dataTransfer.files];
              for (const file of files) void takeFile(file);
            }}
            className={
              "rounded-lg border border-dashed px-4 py-5 " +
              (dragOver ? "border-accent bg-surface-2" : "border-line bg-surface")
            }
          >
            <Upload className="size-5 text-accent" aria-hidden="true" />
            <p className="mt-3 text-sm font-semibold">Suelta el .docx de imprenta</p>
            <p className="mt-1 text-sm leading-snug text-muted">
              Se aceptan los cambios, se quitan comentarios y se dejan las notas al pie. Calibri, Cambria, Arial, Times, Courier, Tahoma, Candara, MS Gothic, Symbol y Wingdings ya están en la app.
            </p>
            <div className="mt-4 flex flex-wrap gap-2">
              <button
                type="button"
                disabled={busy}
                onClick={() => inputRef.current?.click()}
                className="rounded-lg bg-accent px-4 py-2.5 text-sm font-semibold text-bg disabled:opacity-50"
              >
                Abrir Word
              </button>
              <button
                type="button"
                disabled={busy}
                onClick={() => void openSample()}
                className="rounded-lg border border-line bg-surface-2 px-4 py-2.5 text-sm font-semibold text-fg disabled:opacity-50"
              >
                Ver muestra
              </button>
              <button
                type="button"
                disabled={busy}
                onClick={() => fontRef.current?.click()}
                className="rounded-lg border border-line bg-surface-2 px-4 py-2.5 text-sm font-semibold text-fg disabled:opacity-50"
              >
                Subir fuentes
              </button>
            </div>
            <input
              ref={inputRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              className="sr-only"
              onChange={(event) => {
                const file = event.target.files?.[0];
                event.target.value = "";
                if (file) void takeFile(file);
              }}
            />
            <input
              ref={fontRef}
              type="file"
              accept=".ttf,.otf,.ttc,.otc,font/ttf,font/otf"
              multiple
              className="sr-only"
              onChange={(event) => {
                const files = [...(event.target.files ?? [])];
                event.target.value = "";
                if (files.length) void addFonts(files);
              }}
            />
          </div>

          <div className="grid gap-2">
            <Toggle
              pressed={options.acceptRevisions}
              onClick={() => patchOptions({ acceptRevisions: !options.acceptRevisions })}
              label="Quitar marcas de revisión"
              hint="El texto insertado se queda. Lo tachado desaparece. Sin globos ni rayas."
            />
            <Toggle
              pressed={options.removeComments}
              onClick={() => patchOptions({ removeComments: !options.removeComments })}
              label="Quitar comentarios"
              hint="Las notas de revisión salen. Las notas al pie se quedan siempre."
            />
            <Toggle
              pressed={options.keepEndnotes}
              onClick={() => patchOptions({ keepEndnotes: !options.keepEndnotes })}
              label="Conservar notas al final"
              hint="Distintas de los comentarios. Las notas al pie no se pueden apagar."
            />
            <Toggle
              pressed={options.removeHighlights}
              onClick={() => patchOptions({ removeHighlights: !options.removeHighlights })}
              label="Quitar resaltados"
              hint="El amarillo de marcador. No es una marca de revisión; por defecto se respeta."
            />
            <Toggle
              pressed={cropMarks}
              onClick={() => setCropMarks((value) => !value)}
              label="Marcas de corte"
              hint="Solo para imprenta offset. IngramSpark / Lightning Source no las acepta: dejalas apagadas."
            />
            <Toggle
              pressed={grayscale}
              onClick={() => setGrayscale((value) => !value)}
              label="Interior en blanco y negro"
              hint="Para IngramSpark / Lightning Source en B/N: las imágenes PNG pasan a escala de grises sin perder calidad y el PDF sale PDF/X-1a. Los JPEG no se tocan."
            />
            <Toggle
              pressed={substitutes}
              onClick={() => void toggleSubstitutes()}
              label="Completar las fuentes"
              hint="Encendido. Calibri, Cambria, Arial, Times, Courier, Tahoma, Candara, MS Gothic, Symbol y Wingdings ya están en la app. El PDF incrusta ese archivo completo, no el recorte del Word."
            />
            <Toggle
              pressed={fullFont}
              onClick={() => setFullFont((value) => !value)}
              label="Incrustar la fuente entera"
              hint="Encendido. El PDF no se queda con el recorte del Word: mete el archivo de la fuente completo."
            />
          </div>

          {report ? (
            <dl className="grid gap-2 text-sm">
              <Stat icon={<Scissors className="size-4" />} label="Comentarios fuera" value={String(report.commentsRemoved)} />
              <Stat icon={<Type className="size-4" />} label="Inserciones aceptadas" value={String(report.insertionsAccepted)} />
              <Stat icon={<Type className="size-4" />} label="Supresiones quitadas" value={String(report.deletionsRemoved)} />
              <Stat icon={<Pilcrow className="size-4" />} label="Notas al pie" value={String(report.footnotesKept)} />
              <Stat icon={<Pilcrow className="size-4" />} label="Notas al final" value={String(report.endnotesKept)} />
              <Stat
                icon={<ImageIcon className="size-4" />}
                label="Imágenes sin recomprimir"
                value={`${report.images.length} · ${formatBytes(imageBytes)}`}
              />
            </dl>
          ) : null}

          {report && report.fonts.length > 0 ? (
            <div className="grid gap-2">
              <p className="text-sm font-semibold">Fuentes de este Word</p>
              <ul className="grid gap-1 text-sm text-muted">
                {report.fonts.map((font) => {
                  const uploaded = shelf.filter((face) =>
                    fontAliases([face], report).some((alias) => alias.family === font.name),
                  );
                  const covered = uploaded.length > 0 || (substitutes && Boolean(font.substitute));
                  const status = uploaded.length
                    ? `subida · ${uploaded.map(faceLabel).join(", ")} · entera en el PDF`
                    : substitutes && font.substitute
                      ? "ya está en la app"
                      : font.embedded
                        ? "recorte del Word"
                        : "falta · súbela";
                  return (
                    <li key={font.name} className="flex items-baseline justify-between gap-3">
                      <span className="text-fg">{font.name}</span>
                      <span className={covered ? "text-right" : "text-right text-accent"}>{status}</span>
                    </li>
                  );
                })}
              </ul>
              {report.fonts.some((font) => {
                const uploaded = shelf.some((face) =>
                  fontAliases([face], report).some((alias) => alias.family === font.name),
                );
                return !uploaded && !(substitutes && font.substitute);
              }) ? (
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => fontRef.current?.click()}
                  className="rounded-lg border border-accent px-3 py-2.5 text-sm font-semibold text-fg disabled:opacity-40"
                >
                  Subir las que faltan
                </button>
              ) : (
                <p className="text-sm text-muted">Esas fuentes van enteras dentro del PDF.</p>
              )}
            </div>
          ) : null}

          {shelf.length > 0 ? (
            <ul className="grid gap-1 text-sm">
              {shelf.map((face) => {
                const usedAs = report
                  ? fontAliases([face], report)
                      .map((alias) => alias.family)
                      .filter((family) => fontKey(family) !== fontKey(face.family))
                  : [];
                return (
                  <li key={face.id} className="flex items-center justify-between gap-2 text-muted">
                    <span className="text-fg">
                      {face.family} · {faceLabel(face)}
                      {usedAs.length ? ` · como ${usedAs.join(", ")}` : ""}
                    </span>
                    <button type="button" onClick={() => void removeUploaded(face.id)} className="text-muted">
                      Quitar
                    </button>
                  </li>
                );
              })}
            </ul>
          ) : null}

          {imageChecks.length > 0 ? (
            <ul className="grid gap-1 text-sm text-muted">
              {imageChecks.map((image) => (
                <li key={`${image.label}-${image.pixels}`} className="flex items-baseline justify-between gap-3">
                  <span className="text-fg">{image.pixels}</span>
                  <span className={image.ppi > 0 && image.ppi < 300 ? "text-accent" : ""}>
                    {image.ppi > 0 ? `${image.ppi} ppp` : "sin medida"}
                    {image.ppi > 0 && image.ppi < 300 ? " · baja para offset" : " · sin recomprimir"}
                  </span>
                </li>
              ))}
            </ul>
          ) : null}

          {report?.warnings.length ? (
            <ul className="grid gap-2">
              {report.warnings.map((warning) => (
                <li key={warning} className="flex gap-2 text-sm text-accent">
                  <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                  <span>{warning}</span>
                </li>
              ))}
            </ul>
          ) : null}

          {tooTall.length ? (
            <p className="flex gap-2 text-sm text-accent">
              <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
              <span>
                {tooTall.length === 1 ? "La página" : "Las páginas"} {tooTall.join(", ")}{" "}
                {tooTall.length === 1 ? "no entra" : "no entran"} en la hoja: una tabla o una imagen es más alta que
                la página. En el PDF {tooTall.length === 1 ? "saldría" : "saldrían"} más largas que el corte.
              </span>
            </p>
          ) : null}

          {press ? (
            press.pdfx ? (
              <p className="text-sm text-muted">
                El último PDF salió PDF/X-1a:2001: tamaño de corte declarado, fuentes incrustadas y todo en gris.
              </p>
            ) : (
              <div className="grid gap-1 text-sm text-accent">
                <p className="flex gap-2">
                  <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
                  <span>El último PDF no es PDF/X-1a. Lo que falta:</span>
                </p>
                <ul className="grid list-disc gap-1 pl-6">
                  {press.problems.map((problem) => (
                    <li key={problem}>{problem}</li>
                  ))}
                </ul>
              </div>
            )
          ) : null}

          {error ? <p className="text-sm text-accent">{error}</p> : null}
          {status ? <p className="text-sm text-muted">{status}</p> : null}
        </aside>

        <section ref={frameRef} className="pliego-stage min-h-[70vh] overflow-auto bg-bg px-4 py-6 sm:px-8">
          <div ref={styleRef} />
          {phase === "idle" ? (
            <div className="mx-auto grid min-h-[60vh] max-w-xl place-items-center">
              <div className="relative w-full border border-line bg-surface px-8 py-14 text-center">
                <Corner className="top-3 left-3" />
                <Corner className="top-3 right-3 rotate-90" />
                <Corner className="bottom-3 left-3 -rotate-90" />
                <Corner className="right-3 bottom-3 rotate-180" />
                <p className="font-display text-3xl text-fg">Listo para filmar</p>
                <p className="mx-auto mt-3 max-w-sm text-sm leading-relaxed text-muted">
                  La página sale al tamaño del Word. El texto sigue siendo texto, con la fuente dentro del PDF.
                  Las fotos y los PNG no se vuelven a comprimir.
                </p>
              </div>
            </div>
          ) : null}

          <div
            className="pliego-spacer mx-auto"
            style={{
              width: sheet.width ? sheet.width * scale : undefined,
              height: sheet.height ? sheet.height * scale : undefined,
            }}
          >
            <div
              ref={docRef}
              className="pliego-sheet"
              style={{
                transform: `scale(${scale})`,
                transformOrigin: "top left",
                width: sheet.width || undefined,
              }}
            />
          </div>

          {phase === "ready" ? (
            <p className="no-print mx-auto mt-4 max-w-xl text-center text-sm text-muted">
              {name} · {pageCount} {pageCount === 1 ? "página" : "páginas"}
              {pageMm ? ` · ${pageMm}` : ""}
              {fit ? "" : " · tamaño real"}
            </p>
          ) : null}

          {phase === "ready" ? (
            <div className="no-print mt-3 flex justify-center">
              <button
                type="button"
                onClick={() => setFit((value) => !value)}
                className="rounded-lg border border-line px-3 py-2 text-sm text-muted"
              >
                {fit ? "Ver al 100%" : "Ajustar al ancho"}
              </button>
            </div>
          ) : null}
        </section>
      </div>

      <div className="no-print fixed inset-x-0 bottom-0 z-30 border-t border-line bg-bg/95 px-4 py-3 backdrop-blur">
        <div className="mx-auto flex max-w-3xl items-center gap-2">
          <button
            type="button"
            disabled={phase !== "ready" || busy}
            onClick={() => void downloadPdf()}
            className="flex flex-1 items-center justify-center gap-2 rounded-lg bg-accent px-4 py-3.5 text-base font-semibold text-bg disabled:opacity-40"
          >
            <FileDown className="size-5" aria-hidden="true" />
            Descargar PDF
          </button>
          <button
            type="button"
            disabled={!report || busy}
            onClick={() => void downloadDocx()}
            className="rounded-lg border border-line bg-surface px-4 py-3.5 text-sm font-semibold text-fg disabled:opacity-40"
          >
            <span className="inline-flex items-center gap-2">
              <FileText className="size-4" aria-hidden="true" />
              Word
            </span>
          </button>
        </div>
      </div>
    </main>
  );
}

function faceLabel(face: { weight: number; style: "normal" | "italic" }) {
  if (face.style === "italic" && face.weight >= 600) return "negrita cursiva";
  if (face.style === "italic") return "cursiva";
  if (face.weight >= 600) return "negrita";
  if (face.weight <= 300) return "light";
  return "regular";
}

function Stat({ icon, label, value }: { icon: ReactNode; label: string; value: string }) {
  return (
    <div className="flex items-center justify-between gap-3 rounded-lg border border-line bg-surface px-3 py-2">
      <dt className="flex items-center gap-2 text-muted">
        <span className="text-accent" aria-hidden="true">
          {icon}
        </span>
        {label}
      </dt>
      <dd className="font-semibold text-fg">{value}</dd>
    </div>
  );
}

function Corner({ className }: { className: string }) {
  return (
    <span className={`absolute size-4 border-t border-l border-accent ${className}`} aria-hidden="true" />
  );
}

import type { MediaAsset, ServiceItem, Slide } from "../types/presentation";

export type ImportSlidesProgress = (message: string, current?: number, total?: number) => void;

function createBlankSlide(overrides: Partial<Slide> = {}): Slide {
  return {
    id: crypto.randomUUID(),
    type: "text",
    text: "",
    background: "#000000",
    textAlign: "center",
    fontSize: 56,
    ...overrides,
  };
}

/**
 * Render each page of a PDF to a PNG data URL via pdf.js.
 * Install once:  npm i pdfjs-dist
 */
export async function pdfFileToPageDataUrls(
  file: File,
  options: { scale?: number; onProgress?: ImportSlidesProgress } = {},
): Promise<string[]> {
  const scale = options.scale ?? 2;
  const onProgress = options.onProgress;

  // pdfjs-dist v4+ ESM build
  const pdfjs = await import("pdfjs-dist");
  // Vite-friendly worker URL
  const workerMod = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
  pdfjs.GlobalWorkerOptions.workerSrc = workerMod.default;

  onProgress?.("Opening PDF…");
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjs.getDocument({ data }).promise;
  const pages: string[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    onProgress?.(`Rendering page ${i} of ${pdf.numPages}…`, i, pdf.numPages);
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) throw new Error("Could not create a canvas for PDF rendering.");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport, canvas }).promise;
    pages.push(canvas.toDataURL("image/png"));
  }

  return pages;
}

export interface BuiltImport {
  item: ServiceItem;
  mediaAssets: MediaAsset[];
}

/**
 * Build a service item + media assets from PDF page images.
 * Each page becomes one slide with backgroundMedia set to the page image.
 */
export function buildImportedDeckItem(
  fileName: string,
  pageDataUrls: string[],
): BuiltImport {
  const baseName = fileName.replace(/\.(pdf|pptx?)$/i, "") || "Imported slides";
  const now = new Date().toISOString();
  const mediaAssets: MediaAsset[] = [];
  const slides: Slide[] = pageDataUrls.map((source, index) => {
    const asset: MediaAsset = {
      id: crypto.randomUUID(),
      name: `${fileName} — Slide ${index + 1}`,
      kind: "image",
      source,
      size: Math.round(source.length * 0.75), // approximate decoded size
      addedAt: now,
    };
    mediaAssets.push(asset);
    return createBlankSlide({
      text: "",
      background: "#000000",
      backgroundMedia: asset.source,
      sectionLabel: `Slide ${index + 1}`,
    });
  });

  const item: ServiceItem = {
    id: crypto.randomUUID(),
    title: baseName,
    type: "other",
    slides,
  };

  return { item, mediaAssets };
}

/**
 * Import a PDF (or reject non-PDF with guidance).
 * PowerPoint should be exported as PDF first for exact styling — no LibreOffice.
 */
export async function importDeckFile(
  file: File,
  options: { onProgress?: ImportSlidesProgress; scale?: number } = {},
): Promise<BuiltImport> {
  const name = file.name || "deck.pdf";
  const isPdf = /\.pdf$/i.test(name) || file.type === "application/pdf";
  const isPpt = /\.pptx?$/i.test(name);

  if (isPpt) {
    throw new Error(
      "Import a PDF for exact styling. In PowerPoint: File → Save As → PDF, then import that file.",
    );
  }
  if (!isPdf) {
    throw new Error("Choose a PDF file (export from PowerPoint with File → Save As → PDF).");
  }

  const pages = await pdfFileToPageDataUrls(file, {
    scale: options.scale ?? 2,
    onProgress: options.onProgress,
  });
  if (pages.length === 0) throw new Error("No pages were found in this PDF.");
  return buildImportedDeckItem(name, pages);
}

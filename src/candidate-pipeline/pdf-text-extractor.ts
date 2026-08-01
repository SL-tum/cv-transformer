import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import type { TextItem } from "pdfjs-dist/types/src/display/api.js";
import type { CandidatePdfDocument, CandidatePdfPage, CandidatePdfTextItem } from "./model.js";

export interface ExtractPdfOptions {
  relativeTo?: string;
  rejectTextlessPages?: boolean;
}

export async function extractTextPdf(
  filePath: string,
  options: ExtractPdfOptions = {},
): Promise<CandidatePdfDocument> {
  const bytes = await readFile(filePath);
  const loadingTask = getDocument({ data: new Uint8Array(bytes), disableFontFace: true });
  const pdf = await loadingTask.promise;
  const pages: CandidatePdfPage[] = [];
  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent({
        includeMarkedContent: true,
        disableNormalization: true,
      });
      const sourceItems = content.items.filter((item): item is TextItem => "str" in item);
      const items: CandidatePdfTextItem[] = sourceItems.map((item, index) => ({
        index,
        text: item.str,
        direction: item.dir,
        transform: [...item.transform],
        width: item.width,
        height: item.height,
        fontName: item.fontName,
        hasEol: item.hasEOL,
      }));
      const text = textFromItems(items);
      if ((options.rejectTextlessPages ?? true) && !text.trim())
        throw new Error(`PDF page ${pageNumber} has no extractable text: ${filePath}`);
      pages.push({ pageNumber, width: viewport.width, height: viewport.height, text, items });
      page.cleanup();
    }
  } finally {
    await loadingTask.destroy();
  }
  const relativePath = options.relativeTo
    ? path.relative(options.relativeTo, filePath)
    : path.basename(filePath);
  return {
    fileName: path.basename(filePath),
    relativePath,
    sha256: createHash("sha256").update(bytes).digest("hex"),
    byteLength: bytes.byteLength,
    pageCount: pages.length,
    pages,
    fullText: pages.map((page) => `--- Page ${page.pageNumber} ---\n${page.text}`).join("\n\n"),
  };
}

function textFromItems(items: CandidatePdfTextItem[]): string {
  let output = "";
  for (const item of items) {
    output += item.text;
    if (item.hasEol) output += "\n";
    else if (item.text && !/\s$/u.test(item.text)) output += " ";
  }
  return output.replace(/[ \t]+\n/gu, "\n").trimEnd();
}

import { mkdir, readdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { extractTextPdf } from "./pdf-text-extractor.js";
import type { CandidateGroundTruth } from "./model.js";

export interface LoadCandidateFolderOptions {
  candidateId?: string;
  recursive?: boolean;
  rejectTextlessPages?: boolean;
}

export async function loadCandidateGroundTruth(
  folderPath: string,
  options: LoadCandidateFolderOptions = {},
): Promise<CandidateGroundTruth> {
  const pdfPaths = await findPdfs(folderPath, options.recursive ?? false);
  if (!pdfPaths.length) throw new Error(`No PDF files found in candidate folder: ${folderPath}`);
  const documents = [];
  for (const filePath of pdfPaths)
    documents.push(
      await extractTextPdf(filePath, {
        relativeTo: folderPath,
        rejectTextlessPages: options.rejectTextlessPages ?? true,
      }),
    );
  const totalPages = documents.reduce((sum, document) => sum + document.pageCount, 0);
  const pagesWithText = documents.reduce(
    (sum, document) => sum + document.pages.filter((page) => page.text.trim()).length,
    0,
  );
  const totalTextItems = documents.reduce(
    (sum, document) =>
      sum + document.pages.reduce((pageSum, page) => pageSum + page.items.length, 0),
    0,
  );
  return {
    schemaVersion: "1.0",
    candidateId: options.candidateId ?? path.basename(path.resolve(folderPath)),
    generatedAt: new Date().toISOString(),
    documents,
    fullText: documents
      .map((document) => `===== ${document.relativePath} =====\n${document.fullText}`)
      .join("\n\n"),
    integrity: {
      totalDocuments: documents.length,
      totalPages,
      pagesWithText,
      totalTextItems,
      warnings: [],
    },
  };
}

export async function saveCandidateGroundTruth(
  groundTruth: CandidateGroundTruth,
  outputFolder: string,
  baseName = "candidate-ground-truth",
): Promise<{ jsonPath: string; textPath: string }> {
  await mkdir(outputFolder, { recursive: true });
  const jsonPath = path.join(outputFolder, `${baseName}.json`);
  const textPath = path.join(outputFolder, `${baseName}.txt`);
  await Promise.all([
    writeFile(jsonPath, `${JSON.stringify(groundTruth, null, 2)}\n`, "utf8"),
    writeFile(textPath, `${groundTruth.fullText}\n`, "utf8"),
  ]);
  return { jsonPath, textPath };
}

async function findPdfs(folderPath: string, recursive: boolean): Promise<string[]> {
  const entries = await readdir(folderPath, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const fullPath = path.join(folderPath, entry.name);
    if (entry.isFile() && entry.name.toLowerCase().endsWith(".pdf")) files.push(fullPath);
    else if (recursive && entry.isDirectory()) files.push(...(await findPdfs(fullPath, true)));
  }
  return files.sort((a, b) => a.localeCompare(b));
}

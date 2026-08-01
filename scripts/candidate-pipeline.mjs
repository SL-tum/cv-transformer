import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  Gemini35FlashProvider,
  TEMPLATE_OPERATION_SCHEMA_VERSION,
  executeOperationsAtomically,
  exportDocx,
  exportPptx,
  extractDocxTemplate,
  extractPptxTemplate,
  importDocx,
  importPptx,
  loadCandidateGroundTruth,
  mapCandidateToTemplate,
  saveCandidateGroundTruth,
} from "../dist/index.js";

const [candidateFolder, templatePath, outputFolder = "candidate-output"] = process.argv.slice(2);
if (!candidateFolder || !templatePath) {
  throw new Error(
    "Usage: node --env-file=.env scripts/candidate-pipeline.mjs <candidate-folder> <template.docx|template.pptx> [output-folder]",
  );
}

const groundTruth = await loadCandidateGroundTruth(candidateFolder, { recursive: true });
const saved = await saveCandidateGroundTruth(groundTruth, outputFolder);
const templateBytes = await readFile(templatePath);
const extension = path.extname(templatePath).toLowerCase();
const document =
  extension === ".docx"
    ? importDocx(templateBytes)
    : extension === ".pptx"
      ? importPptx(templateBytes)
      : undefined;
if (!document) throw new Error("Template must be a .docx or .pptx file");
const extracted =
  extension === ".docx"
    ? extractDocxTemplate(document, {
        documentId: path.basename(templatePath),
        revision: 1,
        mode: "hybrid",
      })
    : extractPptxTemplate(document, {
        documentId: path.basename(templatePath),
        revision: 1,
        mode: "hybrid",
      });
const mapping = await mapCandidateToTemplate(
  new Gemini35FlashProvider(),
  document,
  extracted.template,
  extracted.bindings,
  groundTruth,
);
const planPath = path.join(outputFolder, "candidate-field-plan.json");
await writeFile(
  planPath,
  `${JSON.stringify(
    {
      documentId: extracted.template.documentId,
      revision: extracted.template.revision,
      extractionWarnings: extracted.warnings,
      fieldPlan: mapping.fieldPlan,
      preview: mapping.preview,
      llmMetadata: mapping.generation.metadata,
    },
    null,
    2,
  )}\n`,
  "utf8",
);

if (!mapping.preview.valid)
  throw new Error(`Operation preview failed: ${mapping.preview.errors.join("; ")}`);
executeOperationsAtomically(document, extracted.template, extracted.bindings, {
  schemaVersion: TEMPLATE_OPERATION_SCHEMA_VERSION,
  documentId: extracted.template.documentId,
  baseRevision: extracted.template.revision,
  operations: mapping.operations,
});
const documentPath = path.join(outputFolder, `output${extension}`);
await writeFile(documentPath, extension === ".docx" ? exportDocx(document) : exportPptx(document));

console.log(`Ground truth JSON: ${saved.jsonPath}`);
console.log(`Ground truth text: ${saved.textPath}`);
console.log(`Field plan: ${planPath}`);
console.log(`Output document: ${documentPath}`);
console.log(`Preview valid: ${mapping.preview.valid}`);
if (!mapping.preview.valid) console.error(mapping.preview.errors);

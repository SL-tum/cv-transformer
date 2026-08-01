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
  cleanDocxLayoutBeforeExport,
  assessGeneratedOutput,
  runOutputQualityRetryLoop,
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
if (extension !== ".docx" && extension !== ".pptx")
  throw new Error("Template must be a .docx or .pptx file");
const maxAttempts = Number(process.env.CANDIDATE_OUTPUT_MAX_ATTEMPTS ?? 3);
const provider = new Gemini35FlashProvider();
const accepted = await runOutputQualityRetryLoop(
  async (attempt, previousErrors) => {
    console.log(`Output generation attempt ${attempt}/${maxAttempts}`);
    if (previousErrors.length) console.log(`Previous rejection: ${previousErrors.join("; ")}`);
    const document = extension === ".docx" ? importDocx(templateBytes) : importPptx(templateBytes);
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
      provider,
      document,
      extracted.template,
      extracted.bindings,
      groundTruth,
    );
    if (!mapping.preview.valid)
      throw new Error(`Operation preview failed: ${mapping.preview.errors.join("; ")}`);
    executeOperationsAtomically(document, extracted.template, extracted.bindings, {
      schemaVersion: TEMPLATE_OPERATION_SCHEMA_VERSION,
      documentId: extracted.template.documentId,
      baseRevision: extracted.template.revision,
      operations: mapping.operations,
    });
    const cleanup = extension === ".docx" ? cleanDocxLayoutBeforeExport(document) : undefined;
    const bytes = extension === ".docx" ? exportDocx(document) : exportPptx(document);
    return { document, extracted, mapping, cleanup, bytes };
  },
  ({ bytes, mapping }) =>
    assessGeneratedOutput(bytes, extension === ".docx" ? "docx" : "pptx", mapping.operations),
  { maxAttempts },
);
const { extracted, mapping, cleanup, bytes } = accepted.value;
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
      llmCalls: mapping.generations.length,
      llmCallMetadata: mapping.generations.map((generation) => generation.metadata),
      outputQuality: accepted.quality,
      outputAttempts: accepted.attempts,
      outputAttemptHistory: accepted.history,
    },
    null,
    2,
  )}\n`,
  "utf8",
);

const documentPath = path.join(outputFolder, `output${extension}`);
await writeFile(documentPath, bytes);

console.log(`Ground truth JSON: ${saved.jsonPath}`);
console.log(`Ground truth text: ${saved.textPath}`);
console.log(`Field plan: ${planPath}`);
console.log(`Output document: ${documentPath}`);
console.log(`Preview valid: ${mapping.preview.valid}`);
console.log(`Output quality valid: ${accepted.quality.valid}`);
console.log(`Output accepted after attempts: ${accepted.attempts}`);
if (cleanup) {
  console.log(`DOCX blank paragraphs removed: ${cleanup.paragraphsRemoved}`);
  console.log(`DOCX blank table rows removed: ${cleanup.tableRowsRemoved}`);
}
if (!mapping.preview.valid) console.error(mapping.preview.errors);

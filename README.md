# Rich Document Tree

TypeScript domain model for carrying DOCX and PPTX in one package without erasing their different layout semantics.

- Shared canonical primitives for text, DrawingML graphics, colors, measurements, styles, themes, resources, and relationships.
- Separate `FlowDocumentRoot` and `PresentationDocumentRoot` models.
- `NativeStore` and `NativePayload` preserve unsupported OOXML parts and extensions.
- Deterministic source-based node IDs support diffing and incremental editing.
- Basic structural and reference validation is included.

The package now includes the Phase 2 minimal round-trip layers:

- OPC ZIP reader/writer, content types, root/part relationships, target resolution, package validation, and byte-preserving opaque parts.
- DOCX importer/exporter for sections, paragraphs, runs, text, basic formatting, styles, numbering-part preservation, tables, inline images, and headers/footers.
- PPTX importer/exporter for presentations, slides, masters, layouts, shapes, text, pictures, groups, placeholders, and themes.
- Semantic exporters patch mapped XML text nodes and merge them back into the preserved native package. Unsupported XML, extensions, and binary resources are not rebuilt or discarded.

```ts
import { createRichDocument, stableNodeId } from "@document-tree/rdt";

const root = { id: "root", type: "flowDocument", sections: [] } as const;
const document = createRichDocument("document-1", root);
const id = stableNodeId({
  sourceFormat: "docx",
  partUri: "/word/document.xml",
  xmlPath: "/w:document/w:body/w:p[1]",
});
```

Round-trip editing:

```ts
import { readFile, writeFile } from "node:fs/promises";
import { exportDocx, importDocx } from "@document-tree/rdt";

const document = importDocx(await readFile("input.docx"));
const paragraph = document.root.sections[0]?.blocks[0];
if (
  paragraph?.type === "paragraph" &&
  "runs" in paragraph &&
  paragraph.runs[0]?.type === "textRun"
) {
  paragraph.runs[0].text = "Updated text";
}
await writeFile("output.docx", exportDocx(document));
```

Run `npm install`, then `npm run check` and `npm test`.

## LLM Template Layer

Phase 4 keeps OOXML details out of model prompts. Controlled templates use markers such as `rdt:field:profile`, `rdt:list:skills`, `rdt:collection:experience`, and `rdt:prototype:experience-item`.

```ts
import {
  evaluateTemplateQuality,
  executeTemplateOperations,
  exportDocx,
  extractDocxTemplate,
  importDocx,
  serializeTemplateForLlm,
} from "@document-tree/rdt";

const document = importDocx(input);
const { template, bindings } = extractDocxTemplate(document, {
  documentId: "resume-template",
  revision: 1,
  markers: {
    profile: {
      label: "Profile",
      constraints: { hard: { maxCharacters: 400 } },
    },
  },
});

// Safe prompt payload: no part URI, XML path, relationship, or native payload.
const promptJson = serializeTemplateForLlm(template);

executeTemplateOperations(document, template, bindings, {
  documentId: "resume-template",
  revision: 1,
  operations: [{ op: "setText", targetId: "profile", value: "New profile" }],
});

const quality = evaluateTemplateQuality(template, bindings, document);
if (!quality.passed) throw new Error(quality.issues.join("; "));
const output = exportDocx(document);
```

DOCX markers are read from Content Control tags/titles and bookmarks. PPTX markers are read from shape names and alternative text. Prototype internals stay in the binding map and do not appear as normal LLM content.

### Hybrid extraction

`extractDocxTemplate` and `extractPptxTemplate` default to hybrid mode. Explicit markers remain highest priority; unmarked content can be inferred from Office roles, labels, paragraph order, style, and layout. Every inferred node carries confidence and evidence.

```ts
const extracted = extractPptxTemplate(document, {
  mode: "hybrid", // "strict" | "hybrid" | "unmarked"
  acceptConfidence: 0.9,
  reviewConfidence: 0.7,
});

console.log(extracted.extraction?.candidates);
console.log(extracted.extraction?.decisions);
```

Accepted proposals are editable. Proposals in the review band are included as non-editable nodes with warnings. Explicit marker coverage is reported as a metric, while editable binding coverage remains the hard quality gate.

## LLM Provider Adapter

No provider SDK is imported by the template, binding, RDT, mutation, or OOXML layers. An outer adapter converts provider output into a strict versioned envelope and previews it before atomic execution.

```ts
import {
  LlmOperationExecutor,
  buildOperationPrompt,
  parseStructuredOutput,
} from "@document-tree/rdt";

const messages = buildOperationPrompt({
  template,
  userInput: "Rewrite my profile",
  allowedOperations: ["setText", "setList"],
});

// Send `messages` and templateOperationJsonSchema through any provider SDK.
const envelope = parseStructuredOutput(providerResponse);
const executor = new LlmOperationExecutor(document, template, bindings, ["setText", "setList"]);

const preview = executor.preview(envelope);
if (!preview.valid) throw new Error(preview.errors.join("; "));
const result = executor.execute(envelope); // transactional; rolls back on failure
```

### Gemini 3.5 Flash

The built-in provider uses the Gemini REST API and native `fetch`; no Google SDK is imported into the project.

```bash
export GEMINI_API_KEY="YOUR_GEMINI_API_KEY"
```

```ts
import {
  Gemini35FlashProvider,
  LlmOperationExecutor,
  generateOperationPreview,
} from "@document-tree/rdt";

const provider = new Gemini35FlashProvider({
  // Prefer GEMINI_API_KEY. This explicit placeholder may also be replaced locally.
  apiKey: process.env.GEMINI_API_KEY ?? "YOUR_GEMINI_API_KEY",
  model: "gemini-3.5-flash",
});

const generated = await generateOperationPreview(
  provider,
  {
    template,
    userInput: "Rewrite the profile and update the skills.",
    allowedOperations: ["setText", "setList"],
    context: { language: "en", purpose: "resume" },
  },
  document,
  bindings,
);

if (!generated.preview.valid) {
  throw new Error(generated.preview.errors.join("; "));
}

new LlmOperationExecutor(document, template, bindings, ["setText", "setList"]).execute(
  generated.envelope,
);
```

Copy [.env.example](./.env.example) to a local `.env` only if your runtime loads dotenv files. The library itself intentionally does not load `.env`; it reads `process.env.GEMINI_API_KEY`.

# Candidate PDF Ground Truth Pipeline

Candidate resumes use an independent ground-truth pipeline around the existing template extractor. It reads every PDF in a folder recursively, extracts the PDF text layer page by page, stores both loss-auditable text items and plain text, then asks Gemini to map supported source facts to every editable template field. The candidate PDF is treated only as a fact source; each target section receives the complete candidate text plus the template section context. The accepted operations are applied to a fresh copy of the template and exported as DOCX or PPTX.

```bash
npm run build
node --env-file=.env scripts/candidate-pipeline.mjs manual/candidates manual/input.docx manual/candidate-output
```

The same command supports PPTX templates:

```bash
node --env-file=.env scripts/candidate-pipeline.mjs manual/candidates manual/input.pptx manual/candidate-output
```

DOCX table sections with headers but no populated rows are treated as editable collections. For
example, an empty `Work experience` table can produce multiple `appendCollectionItem` operations.
The extractor records each schema field's table-cell index, allowing cloned rows to preserve the
template formatting while filling the correct cells.

Before DOCX export, the pipeline removes redundant blank paragraphs and completely blank table
rows. It always preserves the minimum paragraph structure required by Word inside retained cells.

The candidate pipeline accepts an output only after OPC validation, RDT re-import validation,
operation-content verification, whitespace checks, and blank-table-row checks. A rejected output
is regenerated from the untouched template. The default limit is three attempts and can be changed
with `CANDIDATE_OUTPUT_MAX_ATTEMPTS`. The accepted quality report and complete attempt history are
stored in `candidate-field-plan.json`; an output file is written only after acceptance. Every retry
starts from the untouched template, so a rejected attempt cannot leak mutations into the next one.

```bash
# Optional; defaults to 3.
CANDIDATE_OUTPUT_MAX_ATTEMPTS=5 \
  node --env-file=.env scripts/candidate-pipeline.mjs \
  manual/candidates manual/input.docx manual/candidate-output
```

Outputs:

- `candidate-ground-truth.json`: source hashes, page dimensions, every text item and its layout metadata, page text, and merged text.
- `candidate-ground-truth.txt`: human-readable file/page-separated source text.
- `candidate-field-plan.json`: operations for every editable target, LLM metadata, dry-run result, final quality report, accepted-attempt number, and complete retry history.
- `output.docx` or `output.pptx`: the atomically validated operations applied to a copy of the template.

This phase supports PDFs with a real text layer. A page with no extractable text is rejected instead of being silently omitted. Image-only/scanned PDF OCR is intentionally not enabled yet.

Gemini response bodies are automatically stored in `manual/llm-responses/` using UTC timestamp filenames. Each file contains only the model-returned body. Set `LLM_RESPONSE_DIR` to use another directory, or pass `responseLogDirectory: false` to `Gemini35FlashProvider` to disable recording for a specific provider instance.

## Local files and Git

Local secrets, candidate data, manual fixtures, generated Office files, LLM responses, QA renders,
coverage data, test reports, and runtime logs are excluded by `.gitignore`. In particular, do not
move real candidate resumes or API response logs outside the ignored `manual/`, `outputs/`, or
`logs/` paths unless you intentionally want Git to see them.

The TypeScript tests in `test/` remain tracked because they are project source. Generated test
artifacts belong in `test-results/`, `coverage/`, `outputs/`, or `manual/`; those directories are
ignored. Before pushing, verify the staged files:

```bash
git status --short
git diff --cached --name-only
```

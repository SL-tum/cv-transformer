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
const id = stableNodeId({ sourceFormat: "docx", partUri: "/word/document.xml", xmlPath: "/w:document/w:body/w:p[1]" });
```

Round-trip editing:

```ts
import { readFile, writeFile } from "node:fs/promises";
import { exportDocx, importDocx } from "@document-tree/rdt";

const document = importDocx(await readFile("input.docx"));
const paragraph = document.root.sections[0]?.blocks[0];
if (paragraph?.type === "paragraph" && "runs" in paragraph && paragraph.runs[0]?.type === "textRun") {
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
const executor = new LlmOperationExecutor(
  document,
  template,
  bindings,
  ["setText", "setList"],
);

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

new LlmOperationExecutor(
  document,
  template,
  bindings,
  ["setText", "setList"],
).execute(generated.envelope);
```

Copy [.env.example](./.env.example) to a local `.env` only if your runtime loads dotenv files. The library itself intentionally does not load `.env`; it reads `process.env.GEMINI_API_KEY`.

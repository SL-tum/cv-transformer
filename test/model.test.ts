import assert from "node:assert/strict";
import test from "node:test";
import {
  createRichDocument,
  stableNodeId,
  validateDocument,
  type FlowDocumentRoot,
  type PresentationDocumentRoot,
} from "../src/index.js";

test("stable node ids are deterministic and location-sensitive", () => {
  const source = {
    sourceFormat: "pptx" as const,
    partUri: "/ppt/slides/slide1.xml",
    xmlPath: "/p:sld/p:cSld/p:spTree/p:sp[3]",
    nativeId: "7",
  };
  assert.equal(stableNodeId(source), stableNodeId(source));
  assert.notEqual(stableNodeId(source), stableNodeId({ ...source, nativeId: "8" }));
});

test("factory derives DOCX format and validates a flow root", () => {
  const root: FlowDocumentRoot = {
    id: "root",
    type: "flowDocument",
    sections: [{ id: "section", type: "section", blocks: [], pageProperties: {} }],
  };
  const document = createRichDocument("doc", root);
  assert.equal(document.format, "docx");
  assert.deepEqual(validateDocument(document), []);
});

test("presentation validation catches broken layout references", () => {
  const root: PresentationDocumentRoot = {
    id: "root",
    type: "presentation",
    slideSize: { width: { value: 10, unit: "emu" }, height: { value: 10, unit: "emu" } },
    slideMasters: [],
    slideLayouts: [],
    slides: [{ id: "slide", type: "slide", layoutRef: "missing", shapes: [] }],
  };
  const issues = validateDocument(createRichDocument("deck", root));
  assert.equal(issues[0]?.code, "missing-layout");
});

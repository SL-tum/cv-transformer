import { randomUUID } from "node:crypto";
import type { RichDocument } from "../model/core/document.js";
import type { ParagraphNode, TextRunNode } from "../model/text/index.js";
import type { FlowDocumentRoot, SectionNode } from "../model/word/index.js";
import { addPatch, createPatchPlan, parentElementPath } from "../ooxml/patch-plan.js";

const escapeXml = (value: string) => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
const plan = (document: RichDocument<FlowDocumentRoot>) => document.patchPlan ??= createPatchPlan();

export function splitDocxRun(document: RichDocument<FlowDocumentRoot>, paragraph: ParagraphNode, run: TextRunNode, offset: number): TextRunNode {
  if (offset < 0 || offset > run.text.length) throw new RangeError("Run split offset is out of range");
  if (!run.source?.xmlPath) throw new Error("Splitting requires a source-mapped run");
  const right = structuredClone(run); right.id = `text_${randomUUID()}`; right.text = run.text.slice(offset); delete right.source;
  run.text = run.text.slice(0, offset); const index = paragraph.runs.indexOf(run); if (index < 0) throw new Error("Run is not in the paragraph"); paragraph.runs.splice(index + 1, 0, right);
  addPatch(plan(document), { op: "setText", partUri: run.source.partUri, path: run.source.xmlPath, text: run.text });
  addPatch(plan(document), { op: "insertAfter", partUri: run.source.partUri, path: parentElementPath(run.source.xmlPath), xml: runXml(right) });
  return right;
}

export function insertDocxRun(document: RichDocument<FlowDocumentRoot>, paragraph: ParagraphNode, index: number, text: string, properties: TextRunNode["properties"] = {}): TextRunNode {
  if (index < 0 || index > paragraph.runs.length) throw new RangeError("Run insertion index is out of range");
  const run: TextRunNode = { id: `text_${randomUUID()}`, type: "textRun", text, properties }; const anchor = paragraph.runs[index] ?? paragraph.runs[index - 1];
  if (!anchor?.source?.xmlPath) throw new Error("Insertion requires a source-mapped adjacent run");
  addPatch(plan(document), { op: index < paragraph.runs.length ? "insertBefore" : "insertAfter", partUri: anchor.source.partUri, path: parentElementPath(anchor.source.xmlPath), xml: runXml(run) }); paragraph.runs.splice(index, 0, run); return run;
}

export function deleteDocxRun(document: RichDocument<FlowDocumentRoot>, paragraph: ParagraphNode, run: TextRunNode): void { const index = paragraph.runs.indexOf(run); if (index < 0) throw new Error("Run is not in the paragraph"); if (!run.source?.xmlPath) throw new Error("Deletion requires a source-mapped run"); addPatch(plan(document), { op: "remove", partUri: run.source.partUri, path: parentElementPath(run.source.xmlPath) }); paragraph.runs.splice(index, 1); }

export function insertDocxParagraph(document: RichDocument<FlowDocumentRoot>, section: SectionNode, index: number, text = ""): ParagraphNode { if (index < 0 || index > section.blocks.length) throw new RangeError("Paragraph insertion index is out of range"); const paragraph: ParagraphNode = { id: `p_${randomUUID()}`, type: "paragraph", runs: [{ id: `text_${randomUUID()}`, type: "textRun", text, properties: {} }], properties: {} }; const anchor = section.blocks[index] ?? section.blocks[index - 1]; if (!anchor?.source?.xmlPath) throw new Error("Insertion requires a source-mapped adjacent block"); addPatch(plan(document), { op: index < section.blocks.length ? "insertBefore" : "insertAfter", partUri: anchor.source.partUri, path: anchor.source.xmlPath, xml: paragraphXml(paragraph) }); section.blocks.splice(index, 0, paragraph); return paragraph; }
export function deleteDocxParagraph(document: RichDocument<FlowDocumentRoot>, section: SectionNode, paragraph: ParagraphNode): void { const index = section.blocks.indexOf(paragraph); if (index < 0) throw new Error("Paragraph is not in the section"); if (!paragraph.source?.xmlPath) throw new Error("Deletion requires a source-mapped paragraph"); addPatch(plan(document), { op: "remove", partUri: paragraph.source.partUri, path: paragraph.source.xmlPath }); section.blocks.splice(index, 1); }

function runXml(run: TextRunNode): string { const p = run.properties; const props = [p.bold ? "<w:b/>" : "", p.italic ? "<w:i/>" : "", p.fontSize ? `<w:sz w:val="${Math.round(p.fontSize.value * 2)}"/>` : "", p.color?.type === "rgb" ? `<w:color w:val="${p.color.value}"/>` : ""].join(""); return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t${/^\s|\s$/.test(run.text) ? ' xml:space="preserve"' : ""}>${escapeXml(run.text)}</w:t></w:r>`; }
function paragraphXml(paragraph: ParagraphNode): string { return `<w:p>${paragraph.runs.filter((x): x is TextRunNode => x.type === "textRun").map(runXml).join("")}</w:p>`; }

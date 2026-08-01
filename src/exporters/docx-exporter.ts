import type { RichDocument } from "../model/core/document.js";
import type { InlineNode, ParagraphNode } from "../model/text/index.js";
import type { FlowBlockNode, FlowDocumentRoot } from "../model/word/index.js";
import { nativeStoreToPackage, writeOpcPackage } from "../ooxml/opc/package.js";
import { findByElementPath, parseXml, serializeXml } from "../ooxml/xml.js";
import { applyXmlPatchPlan } from "../ooxml/patch-plan.js";

export function exportDocx(document: RichDocument<FlowDocumentRoot>): Uint8Array {
  if (!document.nativeStore) throw new Error("DOCX export requires the preserved NativeStore");
  const pkg = nativeStoreToPackage(document.nativeStore);
  const byPart = new Map<string, Array<{ path: string; text: string }>>();
  const collectInline = (inline: InlineNode) => {
    if (inline.type === "textRun" && inline.source?.xmlPath) {
      const items = byPart.get(inline.source.partUri) ?? [];
      items.push({ path: inline.source.xmlPath, text: inline.text });
      byPart.set(inline.source.partUri, items);
    }
    if (inline.type === "hyperlink") inline.runs.forEach(collectInline);
  };
  const collectParagraph = (paragraph: ParagraphNode) => paragraph.runs.forEach(collectInline);
  const collectBlock = (block: FlowBlockNode) => {
    if (block.type === "paragraph" && "runs" in block) collectParagraph(block);
    else if (block.type === "table" && "rows" in block)
      block.rows.forEach((row) => row.cells.forEach((cell) => cell.blocks.forEach(collectBlock)));
    else if (
      (block.type === "structuredDocumentTag" || block.type === "customXml") &&
      "blocks" in block
    )
      block.blocks.forEach(collectBlock);
  };
  document.root.sections.forEach((section) => section.blocks.forEach(collectBlock));
  document.root.headers?.forEach((part) => part.blocks.forEach(collectBlock));
  document.root.footers?.forEach((part) => part.blocks.forEach(collectBlock));
  for (const [uri, updates] of byPart) {
    const part = pkg.parts.get(uri);
    if (!part?.xml) continue;
    const dom = parseXml(part.xml);
    for (const update of updates) {
      const element = findByElementPath(dom, update.path);
      if (!element) throw new Error(`Cannot map RDT text back to ${uri}${update.path}`);
      while (element.firstChild) element.removeChild(element.firstChild);
      element.appendChild(dom.createTextNode(update.text));
      if (/^\s|\s$/.test(update.text)) element.setAttribute("xml:space", "preserve");
    }
    part.xml = serializeXml(dom);
  }
  applyXmlPatchPlan(pkg, document.patchPlan);
  return writeOpcPackage(pkg);
}

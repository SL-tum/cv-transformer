import type { RichDocument } from "./model/core/document.js";
import type { BaseNode } from "./model/core/node.js";
import type { SlideElementNode } from "./model/presentation/index.js";
import type { CustomXmlNode, FlowBlockNode, StructuredDocumentTagNode, TableNode } from "./model/word/index.js";

export interface ValidationIssue { code: string; message: string; nodeId?: string; path: string; severity: "error" | "warning" }

export function validateDocument(document: RichDocument): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const ids = new Set<string>();
  const visit = (node: BaseNode, path: string) => {
    if (!node.id) issues.push({ code: "missing-node-id", message: "Every node needs a stable id", path, severity: "error" });
    else if (ids.has(node.id)) issues.push({ code: "duplicate-node-id", message: `Duplicate node id: ${node.id}`, nodeId: node.id, path, severity: "error" });
    else ids.add(node.id);
    node.children?.forEach((child, index) => visit(child, `${path}.children[${index}]`));
  };

  visit(document.root, "root");
  if (document.format === "docx" && document.root.type !== "flowDocument") issues.push({ code: "root-format-mismatch", message: "DOCX requires a flowDocument root", path: "root", severity: "error" });
  if (document.format === "pptx" && document.root.type !== "presentation") issues.push({ code: "root-format-mismatch", message: "PPTX requires a presentation root", path: "root", severity: "error" });

  if (document.root.type === "flowDocument") {
    document.root.sections.forEach((section, si) => {
      visit(section, `root.sections[${si}]`);
      section.blocks.forEach((block, bi) => visitFlowBlock(block, `root.sections[${si}].blocks[${bi}]`, visit));
    });
  } else {
    const masters = new Set(document.root.slideMasters.map((x) => x.id));
    const layouts = new Set(document.root.slideLayouts.map((x) => x.id));
    document.root.slideMasters.forEach((master, i) => { visit(master, `root.slideMasters[${i}]`); visitElements(master.shapes, `root.slideMasters[${i}].shapes`, visit); });
    document.root.slideLayouts.forEach((layout, i) => { visit(layout, `root.slideLayouts[${i}]`); visitElements(layout.shapes, `root.slideLayouts[${i}].shapes`, visit); if (layout.masterRef && !masters.has(layout.masterRef)) issues.push({ code: "missing-master", message: `Unknown master: ${layout.masterRef}`, nodeId: layout.id, path: `root.slideLayouts[${i}].masterRef`, severity: "error" }); });
    document.root.slides.forEach((slide, i) => { visit(slide, `root.slides[${i}]`); visitElements(slide.shapes, `root.slides[${i}].shapes`, visit); if (slide.layoutRef && !layouts.has(slide.layoutRef)) issues.push({ code: "missing-layout", message: `Unknown layout: ${slide.layoutRef}`, nodeId: slide.id, path: `root.slides[${i}].layoutRef`, severity: "error" }); });
  }
  return issues;
}

function visitElements(elements: SlideElementNode[], path: string, visit: (node: BaseNode, path: string) => void) { elements.forEach((element, i) => visit(element, `${path}[${i}]`)); }
function visitFlowBlock(block: FlowBlockNode, path: string, visit: (node: BaseNode, path: string) => void) {
  visit(block, path);
  if (isTable(block)) block.rows.forEach((row, ri) => { visit(row, `${path}.rows[${ri}]`); row.cells.forEach((cell, ci) => { visit(cell, `${path}.rows[${ri}].cells[${ci}]`); cell.blocks.forEach((nested, ni) => visitFlowBlock(nested, `${path}.rows[${ri}].cells[${ci}].blocks[${ni}]`, visit)); }); });
  if (hasNestedBlocks(block)) block.blocks.forEach((nested, i) => visitFlowBlock(nested, `${path}.blocks[${i}]`, visit));
}

function isTable(block: FlowBlockNode): block is TableNode { return block.type === "table" && "rows" in block; }
function hasNestedBlocks(block: FlowBlockNode): block is StructuredDocumentTagNode | CustomXmlNode { return (block.type === "structuredDocumentTag" || block.type === "customXml") && "blocks" in block; }

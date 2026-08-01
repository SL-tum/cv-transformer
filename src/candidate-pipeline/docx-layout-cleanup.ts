import type { RichDocument } from "../model/core/document.js";
import type { ParagraphNode } from "../model/text/index.js";
import type {
  FlowBlockNode,
  FlowDocumentRoot,
  TableCellNode,
  TableNode,
} from "../model/word/index.js";
import { addPatch, createPatchPlan } from "../ooxml/patch-plan.js";

export interface DocxLayoutCleanupResult {
  paragraphsRemoved: number;
  tableRowsRemoved: number;
}

/** Removes redundant blank paragraphs without ever leaving a Word table cell empty. */
export function cleanDocxLayoutBeforeExport(
  document: RichDocument<FlowDocumentRoot>,
): DocxLayoutCleanupResult {
  let paragraphsRemoved = 0;
  let tableRowsRemoved = 0;
  const removeParagraph = (paragraph: ParagraphNode): void => {
    if (!paragraph.source?.xmlPath) return;
    addPatch((document.patchPlan ??= createPatchPlan()), {
      op: "remove",
      partUri: paragraph.source.partUri,
      path: paragraph.source.xmlPath,
    });
    paragraphsRemoved++;
  };

  const cleanCell = (cell: TableCellNode): void => {
    cell.blocks.forEach(cleanNestedBlock);
    const paragraphs = cell.blocks.filter(isParagraph);
    const nonBlank = paragraphs.filter((paragraph) => !isBlankParagraph(paragraph));
    const removable = nonBlank.length
      ? paragraphs.filter(isBlankParagraph)
      : paragraphs.slice(0, -1);
    const removeIds = new Set(removable.map((paragraph) => paragraph.id));
    removable.forEach(removeParagraph);
    cell.blocks = cell.blocks.filter((block) => !removeIds.has(block.id));
  };

  const cleanTable = (table: TableNode): void => {
    table.rows.forEach((row) => row.cells.forEach(cleanCell));
    const emptyRows = table.rows.filter((row) => row.cells.every(isBlankCell));
    const removeIds = new Set(emptyRows.map((row) => row.id));
    for (const row of emptyRows) {
      if (!row.source?.xmlPath) continue;
      addPatch((document.patchPlan ??= createPatchPlan()), {
        op: "remove",
        partUri: row.source.partUri,
        path: row.source.xmlPath,
      });
      tableRowsRemoved++;
    }
    table.rows = table.rows.filter((row) => !removeIds.has(row.id));
  };

  function cleanNestedBlock(block: FlowBlockNode): void {
    if (block.type === "table" && "rows" in block) cleanTable(block);
    else if (
      (block.type === "structuredDocumentTag" || block.type === "customXml") &&
      "blocks" in block
    ) {
      block.blocks.forEach(cleanNestedBlock);
    }
  }

  const cleanTopLevel = (blocks: FlowBlockNode[]): void => {
    blocks.forEach(cleanNestedBlock);
    let previousBlank = false;
    const kept: FlowBlockNode[] = [];
    for (const block of blocks) {
      const blank = isParagraph(block) && isBlankParagraph(block);
      if (blank && previousBlank) {
        removeParagraph(block);
        continue;
      }
      kept.push(block);
      previousBlank = blank;
    }
    blocks.splice(0, blocks.length, ...kept);
  };

  document.root.sections.forEach((section) => cleanTopLevel(section.blocks));
  document.root.headers?.forEach((part) => cleanTopLevel(part.blocks));
  document.root.footers?.forEach((part) => cleanTopLevel(part.blocks));
  return { paragraphsRemoved, tableRowsRemoved };
}

function isBlankCell(cell: TableCellNode): boolean {
  return cell.blocks.every((block) => isParagraph(block) && isBlankParagraph(block));
}

function isParagraph(block: FlowBlockNode): block is ParagraphNode {
  return block.type === "paragraph" && "runs" in block;
}

function isBlankParagraph(paragraph: ParagraphNode): boolean {
  return paragraph.runs.every((run) => run.type === "textRun" && run.text.trim().length === 0);
}

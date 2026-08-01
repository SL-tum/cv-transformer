import type { RichDocument } from "../../model/core/document.js";
import type { ParagraphNode, TextRunNode } from "../../model/text/index.js";
import type {
  FlowBlockNode,
  FlowDocumentRoot,
  TableCellNode,
  TableNode,
  TableRowNode,
} from "../../model/word/index.js";
import type { TemplateBinding } from "../bindings/binding.js";
import type {
  CollectionNode,
  ContainerNode,
  FieldGroupNode,
  TextNode,
} from "../model/template-node.js";
import type { ExtractionDecision, ExtractionProposal } from "./candidate.js";
import type { ExtractionResult } from "./extractor.js";

/** Table-aware extraction. It never flattens cells into the surrounding paragraph stream. */
export function mergeDocxTableExtraction(
  document: RichDocument<FlowDocumentRoot>,
  result: ExtractionResult,
): ExtractionResult {
  const existing = new Set(result.template.root.children.map((node) => node.id));
  const proposals: ExtractionProposal[] = [];
  const decisions: ExtractionDecision[] = [];
  const visit = (blocks: FlowBlockNode[]) =>
    blocks.forEach((block) => {
      if (block.type === "table" && "rows" in block)
        extractTable(block, result, existing, proposals, decisions);
      else if (
        (block.type === "structuredDocumentTag" || block.type === "customXml") &&
        "blocks" in block
      )
        visit(block.blocks);
    });
  document.root.sections.forEach((section) => visit(section.blocks));
  if (result.extraction) {
    result.extraction.proposals.push(...proposals);
    result.extraction.decisions.push(...decisions);
  }
  return result;
}

function extractTable(
  table: TableNode,
  result: ExtractionResult,
  existing: Set<string>,
  proposals: ExtractionProposal[],
  decisions: ExtractionDecision[],
) {
  for (let index = 0; index < table.rows.length; index++) {
    const row = table.rows[index]!;
    const label = fullWidthLabel(row, table);
    if (!label) continue;
    const next = table.rows[index + 1];
    const id = unique(slug(label), existing);
    if (next && headerCells(next).length >= 2) {
      const headers = headerCells(next).filter((item) => !/^n[°ºo]?$/i.test(item.label));
      const dataRows: TableRowNode[] = [];
      let cursor = index + 2;
      while (cursor < table.rows.length && !fullWidthLabel(table.rows[cursor]!, table))
        dataRows.push(table.rows[cursor++]!);
      const items = dataRows
        .filter(rowText)
        .map((dataRow, itemIndex) => rowItem(id, dataRow, headers, itemIndex));
      const baseMetadata = metadata(0.98, [
        "Full-width Word table row defines a section",
        "Following row defines multiple field labels",
        "Subsequent rows repeat the same table grid",
      ]);
      const node: CollectionNode = {
        id,
        type: "collection",
        label,
        editable: false,
        repeatable: true,
        items,
        prototypeId: `${id}-item`,
        metadata: {
          ...baseMetadata,
          itemSchema: headers.map(({ label }) => ({ key: slug(label), label, type: "text" })),
        },
      };
      result.template.root.children.push(node);
      result.bindings.bindings[id] = rowsBinding(id, [row, next, ...dataRows]);
      const prototype = dataRows[0];
      if (prototype)
        result.bindings.prototypes[node.prototypeId] = {
          id: node.prototypeId,
          collectionId: id,
          rootNodeIds: [prototype.id],
          cloneStrategy: "deepCloneTableRow",
          insertionAnchorNodeId: prototype.id,
          insertionPosition: "after",
          regenerateNodeIds: true,
          regenerateOfficeIds: false,
          preserveStyles: true,
          clearEditableContent: true,
          fieldBindings: Object.fromEntries(
            headers.map(({ label, cellIndex }) => [
              slug(label),
              {
                fieldId: slug(label),
                markerId: slug(label),
                relativeSourceNodeIds: prototype.cells[cellIndex]
                  ? nodeIds(prototype.cells[cellIndex])
                  : [],
              },
            ]),
          ),
          ...(prototype.source?.xmlPath
            ? {
                locations: [
                  {
                    nodeId: prototype.id,
                    partUri: prototype.source.partUri,
                    xmlPath: prototype.source.xmlPath,
                    role: "root",
                  },
                ],
              }
            : {}),
        };
      addDecision(proposals, decisions, id, "collection", label, 0.98, [
        row.id,
        next.id,
        ...dataRows.map((item) => item.id),
      ]);
      index = cursor - 1;
    } else if (next && !rowText(next)) {
      const paragraphs = next.cells.flatMap(cellParagraphs);
      const node: TextNode = {
        id,
        type: "text",
        label,
        editable: true,
        value: "",
        metadata: metadata(0.98, [
          "Full-width Word table row defines a field label",
          "Following full-width blank row is its value area",
        ]),
      };
      result.template.root.children.push(node);
      result.bindings.bindings[id] = paragraphsBinding(id, paragraphs);
      addDecision(proposals, decisions, id, "text", label, 0.98, [row.id, next.id]);
      index++;
    } else {
      const node: ContainerNode = {
        id,
        type: "container",
        label,
        editable: false,
        children: [],
        metadata: metadata(0.9, ["Full-width Word table row defines a section boundary"]),
      };
      result.template.root.children.push(node);
      result.bindings.bindings[id] = rowsBinding(id, [row]);
      addDecision(proposals, decisions, id, "container", label, 0.9, [row.id]);
    }
  }
}

function fullWidthLabel(row: TableRowNode, table: TableNode): string | undefined {
  const full =
    row.cells.length === 1 ||
    Boolean(row.cells[0]?.columnSpan && row.cells[0]!.columnSpan! >= table.grid.columns.length);
  if (!full) return undefined;
  return row.cells.flatMap(cellParagraphs).map(paragraphText).find(Boolean);
}
function headerCells(row: TableRowNode) {
  return row.cells
    .map((cell, cellIndex) => ({ cellIndex, label: cellText(cell) }))
    .filter((item) => item.label);
}
function cellParagraphs(cell: TableCellNode) {
  return cell.blocks.filter(
    (block): block is ParagraphNode => block.type === "paragraph" && "runs" in block,
  );
}
function paragraphText(paragraph: ParagraphNode) {
  return paragraph.runs
    .filter((run): run is TextRunNode => run.type === "textRun")
    .map((run) => run.text)
    .join("")
    .trim();
}
function cellText(cell: TableCellNode) {
  return cellParagraphs(cell).map(paragraphText).filter(Boolean).join(" ");
}
function rowText(row: TableRowNode) {
  return row.cells.map(cellText).join("").trim();
}
function nodeIds(value: unknown): string[] {
  const ids: string[] = [];
  const visit = (item: unknown) => {
    if (!item || typeof item !== "object" || item instanceof Uint8Array) return;
    const record = item as Record<string, unknown>;
    if (typeof record.id === "string" && typeof record.type === "string") ids.push(record.id);
    for (const [key, child] of Object.entries(record))
      if (!["source", "native", "metadata", "style"].includes(key))
        Array.isArray(child) ? child.forEach(visit) : visit(child);
  };
  visit(value);
  return ids;
}
function rowsBinding(id: string, rows: TableRowNode[]): TemplateBinding {
  return {
    templateNodeId: id,
    sourceNodeIds: [...new Set(rows.flatMap(nodeIds))],
    sourceFormat: "docx",
    representation: "tableFields",
    readStrategy: { type: "collectText" },
    writeStrategy: { type: "clonePrototype" },
    formatStrategy: { type: "preserveExisting" },
    locations: rows.flatMap((row) =>
      row.source?.xmlPath
        ? [
            {
              nodeId: row.id,
              partUri: row.source.partUri,
              xmlPath: row.source.xmlPath,
              role: "root" as const,
            },
          ]
        : [],
    ),
  };
}
function paragraphsBinding(id: string, paragraphs: ParagraphNode[]): TemplateBinding {
  return {
    templateNodeId: id,
    sourceNodeIds: [...new Set(paragraphs.flatMap(nodeIds))],
    sourceFormat: "docx",
    representation: "joinedRuns",
    readStrategy: { type: "collectText" },
    writeStrategy: { type: "replaceTextPreservingRuns" },
    formatStrategy: { type: "preserveExisting" },
    locations: paragraphs.flatMap((paragraph) =>
      paragraph.source?.xmlPath
        ? [
            {
              nodeId: paragraph.id,
              partUri: paragraph.source.partUri,
              xmlPath: paragraph.source.xmlPath,
              role: "paragraph" as const,
            },
          ]
        : [],
    ),
  };
}
function rowItem(
  collectionId: string,
  row: TableRowNode,
  headers: Array<{ cellIndex: number; label: string }>,
  index: number,
): FieldGroupNode {
  return {
    id: `${collectionId}.item.${index + 1}`,
    type: "fieldGroup",
    label: `${collectionId} item ${index + 1}`,
    editable: false,
    fields: headers.map(({ cellIndex, label }) => ({
      id: `${collectionId}.item.${index + 1}.${slug(label)}`,
      key: slug(label),
      label,
      type: "text",
      value: row.cells[cellIndex] ? cellText(row.cells[cellIndex]!) : "",
      editable: true,
    })),
  };
}
function metadata(confidence: number, evidence: string[]) {
  return { extraction: { source: "structuralRule", confidence, status: "accepted", evidence } };
}
function addDecision(
  proposals: ExtractionProposal[],
  decisions: ExtractionDecision[],
  id: string,
  type: ExtractionProposal["type"],
  label: string,
  confidence: number,
  candidateIds: string[],
) {
  const proposal: ExtractionProposal = {
    id,
    type,
    label,
    candidateIds,
    confidence,
    source: "structuralRule",
    evidence: [
      {
        signal: "nativeRole",
        description: "Structure derived from Word table rows, columns, spans, and headers",
      },
    ],
  };
  proposals.push(proposal);
  decisions.push({ proposal, status: "accepted" });
}
function slug(value: string) {
  return (
    value
      .toLowerCase()
      .normalize("NFKD")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "section"
  );
}
function unique(base: string, existing: Set<string>) {
  let id = base;
  let count = 2;
  while (existing.has(id)) id = `${base}-${count++}`;
  existing.add(id);
  return id;
}

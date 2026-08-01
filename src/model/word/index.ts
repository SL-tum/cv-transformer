import type { BaseNode, RichNode } from "../core/node.js";
import type { Length } from "../core/primitives.js";
import type { ParagraphNode } from "../text/index.js";
import type { GraphicNode } from "../drawing/index.js";

export interface TableGrid {
  columns: Array<{ width?: Length }>;
}
export interface TableCellNode extends BaseNode<"tableCell"> {
  blocks: FlowBlockNode[];
  columnSpan?: number;
  rowSpan?: number;
  merge?: { horizontal?: "start" | "continue"; vertical?: "start" | "continue" };
}
export interface TableRowNode extends BaseNode<"tableRow"> {
  cells: TableCellNode[];
  repeatHeader?: boolean;
}
export interface TableNode extends BaseNode<"table"> {
  grid: TableGrid;
  rows: TableRowNode[];
  properties: Record<string, unknown>;
}
export interface StructuredDocumentTagNode extends BaseNode<"structuredDocumentTag"> {
  tag?: string;
  blocks: FlowBlockNode[];
}
export interface AltChunkNode extends BaseNode<"altChunk"> {
  relationshipId: string;
}
export interface CustomXmlNode extends BaseNode<"customXml"> {
  blocks: FlowBlockNode[];
}
export interface BlockDrawingNode extends BaseNode<"blockDrawing"> {
  graphic: GraphicNode;
}
export type FlowBlockNode =
  | ParagraphNode
  | TableNode
  | StructuredDocumentTagNode
  | AltChunkNode
  | CustomXmlNode
  | BlockDrawingNode
  | RichNode;

export interface PageProperties {
  width?: Length;
  height?: Length;
  orientation?: "portrait" | "landscape";
  margins?: Record<string, Length>;
  native?: Record<string, unknown>;
}
export interface SectionNode extends BaseNode<"section"> {
  blocks: FlowBlockNode[];
  pageProperties: PageProperties;
  columns?: Record<string, unknown>;
  headers?: string[];
  footers?: string[];
  footnoteSettings?: Record<string, unknown>;
  endnoteSettings?: Record<string, unknown>;
}
export interface HeaderFooterPart extends BaseNode<"header" | "footer"> {
  partUri: string;
  blocks: FlowBlockNode[];
}
export interface FlowDocumentRoot extends BaseNode<"flowDocument"> {
  sections: SectionNode[];
  headers?: HeaderFooterPart[];
  footers?: HeaderFooterPart[];
}

export interface NumberingLevel {
  level: number;
  format: string;
  text: string;
  start?: number;
  properties?: Record<string, unknown>;
}
export interface NumberingDefinition {
  abstractId: string;
  levels: NumberingLevel[];
}
export interface NumberingInstance {
  id: string;
  definitionId: string;
  overrides?: Array<{ level: number; start?: number; definition?: NumberingLevel }>;
}

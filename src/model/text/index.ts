import type { BaseNode } from "../core/node.js";
import type { Color, Length } from "../core/primitives.js";

export interface FontReference { family?: string; theme?: string; fallback?: string[] }
export type UnderlineStyle = "none" | "single" | "double" | "dotted" | "dash" | "wave";

export interface RunProperties {
  fontFamily?: FontReference;
  fontSize?: Length;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strike?: boolean;
  doubleStrike?: boolean;
  color?: Color;
  highlight?: Color;
  baseline?: number;
  superscript?: boolean;
  subscript?: boolean;
  letterSpacing?: Length;
  kerning?: Length;
  capitalization?: "none" | "smallCaps" | "allCaps";
  language?: string;
  textFill?: unknown;
  textOutline?: unknown;
  textEffects?: Record<string, unknown>;
}

export interface TextRunNode extends BaseNode<"textRun"> { text: string; properties: RunProperties }
export interface BreakNode extends BaseNode<"break"> { kind: "line" | "page" | "column" }
export interface TabNode extends BaseNode<"tab"> {}
export interface HyperlinkNode extends BaseNode<"hyperlink"> { target?: string; relationshipId?: string; runs: InlineNode[] }
export interface FieldNode extends BaseNode<"field"> { instruction: string; result?: InlineNode[]; complex?: boolean }
export interface InlineImageNode extends BaseNode<"inlineImage"> { resourceId: string }
export interface MathNode extends BaseNode<"math"> {}
export interface BookmarkNode extends BaseNode<"bookmark"> { name: string; action: "start" | "end" }

export type InlineNode = TextRunNode | BreakNode | TabNode | HyperlinkNode | FieldNode | InlineImageNode | MathNode | BookmarkNode;

export interface ParagraphProperties {
  styleId?: string;
  alignment?: "left" | "center" | "right" | "justify" | "distributed";
  indentation?: Record<string, Length>;
  spacing?: Record<string, Length>;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  list?: { numberingInstanceId: string; level: number };
  native?: Record<string, unknown>;
}

export interface ParagraphNode extends BaseNode<"paragraph"> { runs: InlineNode[]; properties: ParagraphProperties }
export interface TextBodyProperties { verticalAnchor?: string; wrapping?: string; columns?: number; native?: Record<string, unknown> }
export interface TextBody { paragraphs: ParagraphNode[]; bodyProperties?: TextBodyProperties }

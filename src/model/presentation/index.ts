import type { BaseNode } from "../core/node.js";
import type { Size, Transform } from "../core/primitives.js";
import type { GraphicNode, ShapeNode } from "../drawing/index.js";

export interface PlaceholderNode extends Omit<ShapeNode, "type"> {
  type: "placeholder";
  placeholderType:
    | "title"
    | "body"
    | "centerTitle"
    | "subtitle"
    | "date"
    | "footer"
    | "slideNumber"
    | "picture"
    | "chart"
    | "table"
    | "media"
    | "object";
  placeholderIndex?: number;
  inheritedFrom?: string;
}
export interface MediaNode extends BaseNode<"media"> {
  resourceId: string;
  transform: Transform;
}
export interface OleObjectNode extends BaseNode<"oleObject"> {
  resourceId: string;
  transform: Transform;
  previewResourceId?: string;
}
export interface ContentPartNode extends BaseNode<"contentPart"> {
  relationshipId: string;
  transform?: Transform;
}
export type SlideElementNode =
  GraphicNode | PlaceholderNode | MediaNode | OleObjectNode | ContentPartNode;

export interface TimingNode extends BaseNode {
  durationMs?: number;
  delayMs?: number;
  targetId?: string;
  trigger?: string;
  children?: TimingNode[];
}
export interface TimingTree {
  root: TimingNode;
}
export interface SlideTransition {
  type: string;
  duration?: number;
  advanceOnClick?: boolean;
  advanceAfter?: number;
  soundResourceId?: string;
  nativeProperties?: Record<string, unknown>;
}
export interface NotesNode extends BaseNode<"notes"> {
  elements: SlideElementNode[];
}
export interface SlideNode extends BaseNode<"slide"> {
  layoutRef?: string;
  masterRef?: string;
  showMasterShapes?: boolean;
  followMasterBackground?: boolean;
  background?: Record<string, unknown>;
  shapes: SlideElementNode[];
  transition?: SlideTransition;
  timing?: TimingTree;
  comments?: Array<Record<string, unknown>>;
  notes?: NotesNode;
  hidden?: boolean;
}
export interface SlideLayoutNode extends BaseNode<"slideLayout"> {
  masterRef?: string;
  matchingName?: string;
  shapes: SlideElementNode[];
}
export interface SlideMasterNode extends BaseNode<"slideMaster"> {
  themeRef?: string;
  shapes: SlideElementNode[];
  layoutRefs: string[];
}
export interface NotesMasterNode extends BaseNode<"notesMaster"> {
  shapes: SlideElementNode[];
}
export interface HandoutMasterNode extends BaseNode<"handoutMaster"> {
  shapes: SlideElementNode[];
}
export interface PresentationDocumentRoot extends BaseNode<"presentation"> {
  slideSize: Size;
  slideMasters: SlideMasterNode[];
  slideLayouts: SlideLayoutNode[];
  slides: SlideNode[];
  notesMasters?: NotesMasterNode[];
  handoutMaster?: HandoutMasterNode;
}

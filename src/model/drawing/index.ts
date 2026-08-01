import type { BaseNode } from "../core/node.js";
import type { Color, Transform } from "../core/primitives.js";
import type { TextBody } from "../text/index.js";

export interface Crop { left?: number; top?: number; right?: number; bottom?: number }
export interface GradientStop { position: number; color: Color }
export type Fill =
  | { type: "none" }
  | { type: "solid"; color: Color }
  | { type: "gradient"; stops: GradientStop[]; geometry: Record<string, unknown> }
  | { type: "image"; resourceId: string; crop?: Crop; tile?: Record<string, unknown> }
  | { type: "pattern"; pattern: string; foreground: Color; background: Color }
  | { type: "group" };

export interface LineStyle { width?: number; fill?: Fill; dash?: string; headEnd?: string; tailEnd?: string }
export interface ShapeGeometry { preset?: string; custom?: Record<string, unknown>; adjustments?: Record<string, number> }
export interface ShapeNode extends BaseNode<"shape"> { geometry: ShapeGeometry; transform: Transform; fill?: Fill; line?: LineStyle; effects?: Record<string, unknown>; textBody?: TextBody }
export interface PictureNode extends BaseNode<"picture"> { resourceId: string; transform: Transform; crop?: Crop; effects?: Record<string, unknown> }
export interface GroupNode extends BaseNode<"group"> { transform: Transform; elements: GraphicNode[] }
export interface ConnectorNode extends BaseNode<"connector"> { transform: Transform; startRef?: string; endRef?: string; line?: LineStyle }
export interface ChartNode extends BaseNode<"chart"> { resourceId: string; transform: Transform }
export interface DiagramNode extends BaseNode<"diagram"> { resourceId: string; transform: Transform }
export interface GraphicFrameNode extends BaseNode<"graphicFrame"> { transform: Transform; contentType: string; resourceId?: string }
export type GraphicNode = ShapeNode | PictureNode | GroupNode | ConnectorNode | ChartNode | DiagramNode | GraphicFrameNode;

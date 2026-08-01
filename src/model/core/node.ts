import type { DocumentFormat, Transform } from "./primitives.js";
import type { NativePayload } from "./native.js";

export interface SourceMapping {
  partUri: string;
  xmlPath?: string;
  relationshipId?: string;
  sourceFormat: DocumentFormat;
  nativeId?: string;
}

export interface StyleReference {
  registryId: string;
  role?: string;
}

export interface NodeStyle<T = Record<string, unknown>> {
  references: StyleReference[];
  direct?: Partial<T>;
  computed?: T;
}

export interface BaseNode<TType extends string = string> {
  id: string;
  type: TType;
  children?: RichNode[];
  style?: NodeStyle;
  transform?: Transform;
  metadata?: Record<string, unknown>;
  source?: SourceMapping;
  native?: NativePayload;
}

export interface UnknownNode extends BaseNode<"unknown"> {
  native: NativePayload;
  fallbackRendering?: { resourceId?: string; description?: string };
}

// Open union: built-in nodes are typed in their domains; extensions remain possible.
export type RichNode = BaseNode | UnknownNode;

import type { NativeRelationship } from "./native.js";

export interface StyleDefinition {
  id: string;
  kind: "paragraph" | "character" | "table" | "numbering" | "shape" | "presentation";
  name?: string;
  basedOn?: string;
  next?: string;
  linked?: string;
  properties: Record<string, unknown>;
  native?: Record<string, unknown>;
}

export interface StyleRegistry {
  documentDefaults?: Record<string, unknown>;
  styles: Record<string, StyleDefinition>;
  latentStyles?: Record<string, Record<string, unknown>>;
}

export interface ThemeDefinition {
  id: string;
  name?: string;
  colors: Record<string, unknown>;
  fonts: Record<string, unknown>;
  formats?: Record<string, unknown>;
}

export interface ThemeRegistry { themes: Record<string, ThemeDefinition> }

export interface ResourceBase { id: string; contentType: string; partUri?: string; data?: Uint8Array }
export interface ImageResource extends ResourceBase { width?: number; height?: number; altText?: string }
export interface MediaResource extends ResourceBase { durationMs?: number }
export interface FontResource extends ResourceBase { family?: string }
export interface EmbeddedResource extends ResourceBase { fileName?: string }
export interface StructuredResource extends ResourceBase { xml?: string; workbookResourceId?: string }

export interface ResourceRegistry {
  images: Record<string, ImageResource>;
  media: Record<string, MediaResource>;
  fonts: Record<string, FontResource>;
  embeddedFiles: Record<string, EmbeddedResource>;
  charts: Record<string, StructuredResource>;
  diagrams: Record<string, StructuredResource>;
}

export interface RelationshipRegistry {
  bySource: Record<string, NativeRelationship[]>;
}

export const emptyRegistries = () => ({
  styles: { styles: {} } satisfies StyleRegistry,
  themes: { themes: {} } satisfies ThemeRegistry,
  resources: {
    images: {}, media: {}, fonts: {}, embeddedFiles: {}, charts: {}, diagrams: {},
  } satisfies ResourceRegistry,
  relationships: { bySource: {} } satisfies RelationshipRegistry,
});

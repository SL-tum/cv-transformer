export type MarkerKind = "block" | "field" | "list" | "group" | "collection" | "prototype" | "fixed";
export type ExtractionSource = "explicitMarker" | "structuralInference";
export interface TemplateMarker { raw: string; kind: MarkerKind; id: string; source: ExtractionSource; label?: string; partUri: string; xmlPath: string; rootNodeId?: string; textNodeIds: string[]; textPaths: string[]; nativeXml?: string }
const MARKER = /^rdt:(block|field|list|group|collection|prototype|fixed):([A-Za-z0-9][A-Za-z0-9._-]*)$/;
export function parseMarker(value: string): { kind: MarkerKind; id: string } | undefined { const match = value.trim().match(MARKER); return match?.[1] && match[2] ? { kind: match[1] as MarkerKind, id: match[2] } : undefined; }

import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import type { ContentConstraints } from "../model/constraints.js";
import type { LlmTemplateDocument } from "../model/template-document.js";
import type { TemplateBindingMap } from "../bindings/binding-map.js";
import type { CandidateExtractionResult } from "./candidate.js";

export interface MarkerConfiguration {
  label?: string;
  editable?: boolean;
  required?: boolean;
  constraints?: ContentConstraints;
  prototypeId?: string;
  collectionId?: string;
  fields?: Array<{ id: string; key: string; markerId: string; type?: "text" | "list" }>;
}
export type ExtractionMode = "strict" | "hybrid" | "unmarked";
export interface ExtractionOptions {
  documentId?: string;
  revision?: number;
  markers?: Record<string, MarkerConfiguration>;
  requireExplicitMarkers?: boolean;
  mode?: ExtractionMode;
  acceptConfidence?: number;
  reviewConfidence?: number;
}
export interface ExtractionResult {
  template: LlmTemplateDocument;
  bindings: TemplateBindingMap;
  warnings: string[];
  extraction?: CandidateExtractionResult;
}
export type AnyRichDocument = RichDocument<DocumentRoot>;

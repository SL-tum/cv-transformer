import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import type { ContentConstraints } from "../model/constraints.js";
import type { LlmTemplateDocument } from "../model/template-document.js";
import type { TemplateBindingMap } from "../bindings/binding-map.js";

export interface MarkerConfiguration { label?: string; editable?: boolean; required?: boolean; constraints?: ContentConstraints; prototypeId?: string; collectionId?: string; fields?: Array<{ id: string; key: string; markerId: string; type?: "text" | "list" }> }
export interface ExtractionOptions { documentId?: string; revision?: number; markers?: Record<string, MarkerConfiguration>; requireExplicitMarkers?: boolean }
export interface ExtractionResult { template: LlmTemplateDocument; bindings: TemplateBindingMap; warnings: string[] }
export type AnyRichDocument = RichDocument<DocumentRoot>;

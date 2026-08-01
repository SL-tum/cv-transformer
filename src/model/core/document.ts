import type { NativeStore } from "./native.js";
import { emptyRegistries, type RelationshipRegistry, type ResourceRegistry, type StyleRegistry, type ThemeRegistry } from "./registries.js";
import type { FlowDocumentRoot } from "../word/index.js";
import type { PresentationDocumentRoot } from "../presentation/index.js";
import type { XmlPatchPlan } from "../../ooxml/patch-plan.js";

export type DocumentRoot = FlowDocumentRoot | PresentationDocumentRoot;
export interface DocumentMetadata { title?: string; subject?: string; creator?: string; keywords?: string[]; createdAt?: string; modifiedAt?: string; custom?: Record<string, unknown> }
export interface CompatibilityInfo { application?: string; applicationVersion?: string; conformance?: "strict" | "transitional"; features?: string[] }

export interface RichDocument<TRoot extends DocumentRoot = DocumentRoot> {
  id: string;
  format: TRoot extends FlowDocumentRoot ? "docx" : TRoot extends PresentationDocumentRoot ? "pptx" : "docx" | "pptx";
  metadata: DocumentMetadata;
  root: TRoot;
  styles: StyleRegistry;
  themes: ThemeRegistry;
  resources: ResourceRegistry;
  relationships: RelationshipRegistry;
  nativeStore?: NativeStore;
  compatibility: CompatibilityInfo;
  patchPlan?: XmlPatchPlan;
}

export function createRichDocument<TRoot extends DocumentRoot>(id: string, root: TRoot): RichDocument<TRoot> {
  const registries = emptyRegistries();
  const format = (root.type === "flowDocument" ? "docx" : "pptx") as RichDocument<TRoot>["format"];
  return { id, format, metadata: {}, root, ...registries, compatibility: {} };
}

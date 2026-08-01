import type { PrototypeBinding, TemplateBinding } from "./binding.js";
export interface TemplateBindingMap {
  documentId: string;
  templateRevision: number;
  bindings: Record<string, TemplateBinding>;
  prototypes: Record<string, PrototypeBinding>;
}

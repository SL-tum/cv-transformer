import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import { indexRdtNodes } from "../markers/common.js";
import type { TemplateBindingMap } from "./binding-map.js";
import type { LlmTemplateDocument } from "../model/template-document.js";
import type { TemplateNode } from "../model/template-node.js";

export interface BindingValidationIssue {
  code: string;
  message: string;
  targetId?: string;
}
export function validateBindingMap(
  template: LlmTemplateDocument,
  map: TemplateBindingMap,
  document: RichDocument<DocumentRoot>,
): BindingValidationIssue[] {
  const issues: BindingValidationIssue[] = [];
  const nodes = flatten(template.root);
  const rdt = indexRdtNodes(document).byId;
  const ids = new Set<string>();
  for (const node of nodes) {
    if (ids.has(node.id))
      issues.push({
        code: "duplicate-template-id",
        message: `Duplicate template id: ${node.id}`,
        targetId: node.id,
      });
    ids.add(node.id);
    if (node.editable && !map.bindings[node.id])
      issues.push({
        code: "unbound-editable-node",
        message: `Editable node is not bound: ${node.id}`,
        targetId: node.id,
      });
    if (node.type === "collection" && (!node.prototypeId || !map.prototypes[node.prototypeId]))
      issues.push({
        code: "missing-prototype",
        message: `Collection has no valid prototype: ${node.id}`,
        targetId: node.id,
      });
  }
  for (const binding of Object.values(map.bindings))
    for (const id of binding.sourceNodeIds)
      if (
        !rdt.has(id) &&
        !binding.locations?.some((location) => location.nodeId === id && location.xmlPath)
      )
        issues.push({
          code: "unresolved-rdt-node",
          message: `Binding ${binding.templateNodeId} references missing RDT node ${id}`,
          targetId: binding.templateNodeId,
        });
  return issues;
}
export function flatten(root: TemplateNode): TemplateNode[] {
  const result = [root];
  if (root.type === "container") root.children.forEach((node) => result.push(...flatten(node)));
  else if (root.type === "collection") root.items.forEach((node) => result.push(...flatten(node)));
  return result;
}

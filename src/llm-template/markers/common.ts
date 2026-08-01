import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import type { BaseNode } from "../../model/core/node.js";

export interface RdtNodeIndex {
  byId: Map<string, BaseNode>;
  bySource: Map<string, BaseNode>;
}
export function indexRdtNodes(document: RichDocument<DocumentRoot>): RdtNodeIndex {
  const byId = new Map<string, BaseNode>();
  const bySource = new Map<string, BaseNode>();
  const seen = new Set<object>();
  const visit = (value: unknown) => {
    if (
      !value ||
      typeof value !== "object" ||
      seen.has(value as object) ||
      value instanceof Uint8Array
    )
      return;
    seen.add(value as object);
    const record = value as Record<string, unknown>;
    if (typeof record.id === "string" && typeof record.type === "string") {
      const node = value as BaseNode;
      byId.set(node.id, node);
      if (node.source?.xmlPath)
        bySource.set(`${node.source.partUri}\u001f${node.source.xmlPath}`, node);
    }
    for (const [key, child] of Object.entries(record))
      if (!["native", "metadata", "style", "source"].includes(key))
        Array.isArray(child) ? child.forEach(visit) : visit(child);
  };
  visit(document.root);
  return { byId, bySource };
}
export const sourceKey = (partUri: string, xmlPath: string) => `${partUri}\u001f${xmlPath}`;

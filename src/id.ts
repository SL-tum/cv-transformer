import { createHash } from "node:crypto";
import type { SourceMapping } from "./model/core/node.js";

/** Deterministic ID derived from native location; stable across repeated imports. */
export function stableNodeId(source: SourceMapping, prefix = "node"): string {
  const identity = [
    source.sourceFormat,
    source.partUri,
    source.xmlPath ?? "",
    source.nativeId ?? "",
    source.relationshipId ?? "",
  ].join("\u001f");
  return `${prefix}_${createHash("sha256").update(identity).digest("hex").slice(0, 20)}`;
}

import type { LlmTemplateDocument } from "../model/template-document.js";
import type { TemplateOperationBatch } from "../model/operations.js";
export const serializeTemplateForLlm = (template: LlmTemplateDocument, space = 2): string =>
  JSON.stringify(template, null, space);
export function parseOperationBatch(json: string): TemplateOperationBatch {
  const value: unknown = JSON.parse(json);
  if (!value || typeof value !== "object") throw new Error("Operation batch must be an object");
  const batch = value as Record<string, unknown>;
  if (
    typeof batch.documentId !== "string" ||
    typeof batch.revision !== "number" ||
    !Array.isArray(batch.operations)
  )
    throw new Error("Operation batch requires documentId, revision, and operations");
  const supported = new Set([
    "setText",
    "setList",
    "appendListItem",
    "removeListItem",
    "appendCollectionItem",
    "updateCollectionItem",
    "removeCollectionItem",
  ]);
  for (const [index, operation] of batch.operations.entries()) {
    if (!operation || typeof operation !== "object")
      throw new Error(`Invalid operation at index ${index}`);
    const item = operation as Record<string, unknown>;
    if (typeof item.op !== "string" || !supported.has(item.op) || typeof item.targetId !== "string")
      throw new Error(`Invalid operation at index ${index}`);
    if ((item.op === "setText" || item.op === "appendListItem") && typeof item.value !== "string")
      throw new Error(`Operation ${index} requires string value`);
    if (
      item.op === "setList" &&
      (!Array.isArray(item.items) || !item.items.every((entry) => typeof entry === "string"))
    )
      throw new Error(`Operation ${index} requires string items`);
    if (
      (item.op === "appendCollectionItem" || item.op === "updateCollectionItem") &&
      (!item.value || typeof item.value !== "object" || Array.isArray(item.value))
    )
      throw new Error(`Operation ${index} requires object value`);
    if (
      (item.op === "updateCollectionItem" || item.op === "removeCollectionItem") &&
      typeof item.itemId !== "string"
    )
      throw new Error(`Operation ${index} requires itemId`);
    if (
      item.op === "removeListItem" &&
      typeof item.itemId !== "string" &&
      typeof item.index !== "number"
    )
      throw new Error(`Operation ${index} requires itemId or index`);
  }
  return value as TemplateOperationBatch;
}

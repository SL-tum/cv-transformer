import type { TemplateOperation } from "../llm-template/index.js";
import {
  TEMPLATE_OPERATION_SCHEMA_VERSION,
  type TemplateOperationEnvelope,
} from "./structured-output.js";

export class StructuredOutputError extends Error {
  constructor(
    message: string,
    public readonly code: string,
  ) {
    super(message);
    this.name = "StructuredOutputError";
  }
}
export function parseStructuredOutput(input: string | unknown): TemplateOperationEnvelope {
  let value: unknown;
  if (typeof input === "string") {
    try {
      value = JSON.parse(input);
    } catch {
      throw new StructuredOutputError("LLM response is not valid JSON", "INVALID_JSON");
    }
  } else value = input;
  const object = record(value, "Envelope");
  exactKeys(object, ["schemaVersion", "documentId", "baseRevision", "operations"], "Envelope");
  if (object.schemaVersion !== TEMPLATE_OPERATION_SCHEMA_VERSION)
    throw new StructuredOutputError("Unsupported schemaVersion", "SCHEMA_VERSION");
  if (typeof object.documentId !== "string" || !object.documentId)
    throw new StructuredOutputError("documentId must be a non-empty string", "SCHEMA");
  if (!Number.isInteger(object.baseRevision) || Number(object.baseRevision) < 0)
    throw new StructuredOutputError("baseRevision must be a non-negative integer", "SCHEMA");
  if (!Array.isArray(object.operations))
    throw new StructuredOutputError("operations must be an array", "SCHEMA");
  return {
    schemaVersion: TEMPLATE_OPERATION_SCHEMA_VERSION,
    documentId: object.documentId,
    baseRevision: Number(object.baseRevision),
    operations: object.operations.map(parseOperation),
  };
}
function parseOperation(value: unknown, index: number): TemplateOperation {
  const item = record(value, `Operation ${index}`);
  if (typeof item.op !== "string" || typeof item.targetId !== "string")
    throw new StructuredOutputError(`Operation ${index} requires op and targetId`, "SCHEMA");
  const base = ["op", "targetId"];
  if (item.op === "setText" || item.op === "appendListItem") {
    exactKeys(item, [...base, "value"], `Operation ${index}`);
    if (typeof item.value !== "string") fail(index, "value must be a string");
    return { op: item.op, targetId: item.targetId, value: item.value };
  }
  if (item.op === "setList") {
    exactKeys(item, [...base, "items"], `Operation ${index}`);
    if (!Array.isArray(item.items) || !item.items.every((x) => typeof x === "string"))
      fail(index, "items must be strings");
    return { op: "setList", targetId: item.targetId, items: item.items as string[] };
  }
  if (item.op === "removeListItem") {
    exactKeys(item, [...base, "itemId", "index"], `Operation ${index}`, true);
    if (typeof item.itemId !== "string" && !Number.isInteger(item.index))
      fail(index, "itemId or index is required");
    return {
      op: "removeListItem",
      targetId: item.targetId,
      ...(typeof item.itemId === "string" ? { itemId: item.itemId } : {}),
      ...(Number.isInteger(item.index) ? { index: Number(item.index) } : {}),
    };
  }
  if (item.op === "appendCollectionItem") {
    exactKeys(item, [...base, "value"], `Operation ${index}`);
    return { op: item.op, targetId: item.targetId, value: collectionValue(item.value, index) };
  }
  if (item.op === "updateCollectionItem") {
    exactKeys(item, [...base, "itemId", "value"], `Operation ${index}`);
    if (typeof item.itemId !== "string") fail(index, "itemId is required");
    return {
      op: item.op,
      targetId: item.targetId,
      itemId: item.itemId as string,
      value: collectionValue(item.value, index),
    };
  }
  if (item.op === "removeCollectionItem") {
    exactKeys(item, [...base, "itemId"], `Operation ${index}`);
    if (typeof item.itemId !== "string") fail(index, "itemId is required");
    return { op: item.op, targetId: item.targetId, itemId: item.itemId as string };
  }
  throw new StructuredOutputError(`Unknown operation type: ${item.op}`, "UNKNOWN_OPERATION");
}
function collectionValue(value: unknown, index: number): Record<string, string | string[]> {
  const item = record(value, `Operation ${index} value`);
  for (const [key, entry] of Object.entries(item))
    if (
      typeof entry !== "string" &&
      (!Array.isArray(entry) || !entry.every((x) => typeof x === "string"))
    )
      fail(index, `collection field ${key} must be string or string[]`);
  return item as Record<string, string | string[]>;
}
function record(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new StructuredOutputError(`${label} must be an object`, "SCHEMA");
  return value as Record<string, unknown>;
}
function exactKeys(
  value: Record<string, unknown>,
  allowed: string[],
  label: string,
  optional = false,
) {
  const unknown = Object.keys(value).filter((key) => !allowed.includes(key));
  if (unknown.length)
    throw new StructuredOutputError(
      `${label} contains unknown fields: ${unknown.join(", ")}`,
      "UNKNOWN_FIELD",
    );
  if (!optional)
    for (const key of allowed)
      if (!(key in value)) throw new StructuredOutputError(`${label} is missing ${key}`, "SCHEMA");
}
function fail(index: number, message: string): never {
  throw new StructuredOutputError(`Operation ${index}: ${message}`, "SCHEMA");
}

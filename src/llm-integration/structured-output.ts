import type { TemplateOperation } from "../llm-template/index.js";
export const TEMPLATE_OPERATION_SCHEMA_VERSION = "1.0" as const;
export interface TemplateOperationEnvelope {
  schemaVersion: typeof TEMPLATE_OPERATION_SCHEMA_VERSION;
  documentId: string;
  baseRevision: number;
  operations: TemplateOperation[];
}

export const templateOperationJsonSchema = {
  type: "object",
  additionalProperties: false,
  required: ["schemaVersion", "documentId", "baseRevision", "operations"],
  properties: {
    schemaVersion: { const: "1.0" },
    documentId: { type: "string" },
    baseRevision: { type: "integer", minimum: 0 },
    operations: {
      type: "array",
      items: {
        oneOf: [
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "value"],
            properties: {
              op: { const: "setText" },
              targetId: { type: "string" },
              value: { type: "string" },
            },
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "items"],
            properties: {
              op: { const: "setList" },
              targetId: { type: "string" },
              items: { type: "array", items: { type: "string" } },
            },
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "value"],
            properties: {
              op: { const: "appendListItem" },
              targetId: { type: "string" },
              value: { type: "string" },
            },
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId"],
            properties: {
              op: { const: "removeListItem" },
              targetId: { type: "string" },
              itemId: { type: "string" },
              index: { type: "integer", minimum: 0 },
            },
            anyOf: [{ required: ["itemId"] }, { required: ["index"] }],
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "value"],
            properties: {
              op: { const: "appendCollectionItem" },
              targetId: { type: "string" },
              value: { type: "object" },
            },
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "itemId", "value"],
            properties: {
              op: { const: "updateCollectionItem" },
              targetId: { type: "string" },
              itemId: { type: "string" },
              value: { type: "object" },
            },
          },
          {
            type: "object",
            additionalProperties: false,
            required: ["op", "targetId", "itemId"],
            properties: {
              op: { const: "removeCollectionItem" },
              targetId: { type: "string" },
              itemId: { type: "string" },
            },
          },
        ],
      },
    },
  },
} as const;

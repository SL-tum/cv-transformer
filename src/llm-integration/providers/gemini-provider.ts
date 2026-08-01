import { buildOperationPrompt } from "../prompt-builder.js";
import type { LlmGenerationRequest, LlmGenerationResult, LlmProvider } from "../provider.js";
import { parseStructuredOutput } from "../response-parser.js";
import { recordLlmResponseBody } from "../response-recorder.js";
import { templateOperationJsonSchema } from "../structured-output.js";
import type { TemplateNode } from "../../llm-template/index.js";

export const GEMINI_API_KEY_PLACEHOLDER = "YOUR_GEMINI_API_KEY";
export interface GeminiProviderOptions {
  apiKey?: string;
  model?: string;
  baseUrl?: string;
  fetch?: typeof fetch;
  temperature?: number;
  responseLogDirectory?: string | false;
}

export class Gemini35FlashProvider implements LlmProvider {
  readonly model: string;
  private readonly apiKey: string;
  private readonly baseUrl: string;
  private readonly fetcher: typeof fetch;
  private readonly temperature: number;
  private readonly responseLogDirectory: string | false | undefined;
  constructor(options: GeminiProviderOptions = {}) {
    this.apiKey = options.apiKey ?? process.env.GEMINI_API_KEY ?? GEMINI_API_KEY_PLACEHOLDER;
    this.model = options.model ?? "gemini-3.5-flash";
    this.baseUrl = (options.baseUrl ?? "https://generativelanguage.googleapis.com").replace(
      /\/$/,
      "",
    );
    this.fetcher = options.fetch ?? globalThis.fetch;
    this.temperature = options.temperature ?? 0.2;
    this.responseLogDirectory = options.responseLogDirectory;
  }

  async generateOperations(request: LlmGenerationRequest): Promise<LlmGenerationResult> {
    if (!this.apiKey || this.apiKey === GEMINI_API_KEY_PLACEHOLDER)
      throw new Error(
        "Gemini API key is not configured. Set GEMINI_API_KEY or pass apiKey to Gemini35FlashProvider.",
      );
    const messages = request.messages ?? buildOperationPrompt(request);
    const system = messages.find((message) => message.role === "system")?.content ?? "";
    const user = messages
      .filter((message) => message.role === "user")
      .map((message) => message.content)
      .join("\n");
    const response = await this.fetcher(
      `${this.baseUrl}/v1beta/models/${encodeURIComponent(this.model)}:generateContent`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": this.apiKey },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: system }] },
          contents: [{ role: "user", parts: [{ text: user }] }],
          generationConfig: {
            temperature: this.temperature,
            responseMimeType: "application/json",
            responseJsonSchema: responseSchemaForRequest(request),
          },
        }),
      },
    );
    const requestId =
      response.headers.get("x-request-id") ??
      response.headers.get("x-goog-request-id") ??
      undefined;
    const payload: unknown = await response
      .json()
      .catch(async () => ({ raw: await response.text().catch(() => "") }));
    if (!response.ok) throw new GeminiApiError(response.status, geminiErrorMessage(payload));
    const text = responseText(payload);
    if (this.responseLogDirectory !== false) {
      await recordLlmResponseBody(text, this.responseLogDirectory);
    }
    const envelope = parseStructuredOutput(text);
    if (envelope.documentId !== request.template.documentId)
      throw new Error(
        `Gemini returned documentId ${envelope.documentId}; expected ${request.template.documentId}`,
      );
    if (envelope.baseRevision !== request.template.revision)
      throw new Error(
        `Gemini returned baseRevision ${envelope.baseRevision}; expected ${request.template.revision}`,
      );
    for (const operation of envelope.operations)
      if (!request.allowedOperations.includes(operation.op))
        throw new Error(`Gemini returned disallowed operation: ${operation.op}`);
    return {
      operations: envelope.operations,
      metadata: {
        provider: "google-gemini",
        model: this.model,
        ...(requestId ? { requestId } : {}),
      },
    };
  }
}

function responseSchemaForRequest(request: LlmGenerationRequest): unknown {
  const schema = structuredClone(templateOperationJsonSchema) as unknown as {
    properties: {
      operations: {
        items: {
          oneOf: Array<{
            properties: Record<string, Record<string, unknown>>;
          }>;
        };
      };
    };
  };
  const variants = schema.properties.operations.items.oneOf.filter((variant) =>
    request.allowedOperations.includes(String(variant.properties.op?.const)),
  );
  schema.properties.operations.items.oneOf = variants;
  if (request.allowedOperations.includes("appendCollectionItem")) {
    const targetId = record(request.userInput)?.currentTargetId;
    const target =
      typeof targetId === "string" ? findTemplateNode(request.template.root, targetId) : undefined;
    const itemSchema = record(target?.metadata)?.itemSchema;
    if (Array.isArray(itemSchema)) {
      const keys = itemSchema
        .map((item) => record(item)?.key)
        .filter((key): key is string => typeof key === "string");
      const collectionVariant = variants.find(
        (variant) => variant.properties.op?.const === "appendCollectionItem",
      );
      if (collectionVariant && keys.length) {
        collectionVariant.properties.value = {
          type: "object",
          additionalProperties: false,
          required: keys,
          properties: Object.fromEntries(keys.map((key) => [key, { type: "string" }])),
        };
      }
    }
  }
  return schema;
}

function findTemplateNode(node: TemplateNode, id: string): TemplateNode | undefined {
  if (node.id === id) return node;
  const children =
    node.type === "container" ? node.children : node.type === "collection" ? node.items : [];
  for (const child of children) {
    const found = findTemplateNode(child, id);
    if (found) return found;
  }
  return undefined;
}

export class GeminiApiError extends Error {
  constructor(
    public readonly status: number,
    message: string,
  ) {
    super(`Gemini API error ${status}: ${message}`);
    this.name = "GeminiApiError";
  }
}
function responseText(payload: unknown): string {
  const candidate = record(payload)?.candidates;
  if (!Array.isArray(candidate)) throw new Error("Gemini response has no candidates");
  const parts = record(record(candidate[0])?.content)?.parts;
  if (!Array.isArray(parts)) throw new Error("Gemini response has no content parts");
  const text = parts
    .map((part) => record(part)?.text)
    .find((value): value is string => typeof value === "string");
  if (!text) throw new Error("Gemini response contains no text");
  return text;
}
function geminiErrorMessage(payload: unknown): string {
  const object = record(payload);
  const error = record(object?.error);
  return typeof error?.message === "string" ? error.message : JSON.stringify(payload);
}
function record(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}
/** Gemini supports a JSON Schema subset. Convert const/oneOf to supported enum/anyOf forms. */
function geminiCompatibleSchema(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(geminiCompatibleSchema);
  if (!value || typeof value !== "object") return value;
  const source = value as Record<string, unknown>;
  const output: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(source)) {
    if (key === "const") output.enum = [child];
    else if (key === "oneOf") output.anyOf = geminiCompatibleSchema(child);
    else output[key] = geminiCompatibleSchema(child);
  }
  return output;
}

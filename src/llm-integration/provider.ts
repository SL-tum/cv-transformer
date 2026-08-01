import type { LlmTemplateDocument, TemplateOperation } from "../llm-template/index.js";

export interface LlmGenerationRequest {
  template: LlmTemplateDocument;
  userInput: unknown;
  allowedOperations: string[];
  context?: { language?: string; purpose?: string };
}
export interface LlmGenerationResult {
  operations: TemplateOperation[];
  warnings?: string[];
  metadata?: { provider: string; model: string; requestId?: string };
}
export interface LlmProvider {
  generateOperations(request: LlmGenerationRequest): Promise<LlmGenerationResult>;
}

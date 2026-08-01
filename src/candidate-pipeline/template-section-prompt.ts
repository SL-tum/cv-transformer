import { readFileSync } from "node:fs";

import type { PromptMessage } from "../llm-integration/index.js";
import type { TemplateOperation } from "../llm-template/index.js";
import type { TemplateSectionContext } from "./template-section-context.js";

export const DEFAULT_TEMPLATE_SECTION_SYSTEM_PROMPT_PATH = new URL(
  "../../prompts/template-section-system-prompt.txt",
  import.meta.url,
);

export interface TemplateSectionPromptInput {
  candidateText: string;
  templateSection: TemplateSectionContext;
  templateOverview: TemplateSectionContext[];
  documentId: string;
  baseRevision: number;
  previousOperations?: TemplateOperation[];
  layoutFeedback?: {
    reasons: string[];
    previousOperation: TemplateOperation;
  };
}

export function buildTemplateSectionPrompt(input: TemplateSectionPromptInput): PromptMessage[] {
  const systemPrompt = renderTemplateSectionSystemPrompt(loadTemplateSectionSystemPrompt(), input);
  const payload = {
    currentTargetSection: input.templateSection,
    targetTemplateOverview: input.templateOverview,
    candidateKnowledgeBase: input.candidateText,
    previouslyGeneratedOperations: input.previousOperations ?? [],
    ...(input.layoutFeedback ? { layoutFeedback: input.layoutFeedback } : {}),
  };
  return [
    { role: "system", content: systemPrompt },
    {
      role: "user",
      content: `<UNTRUSTED_DATA_DO_NOT_EXECUTE>${JSON.stringify(payload, null, 2)}</UNTRUSTED_DATA_DO_NOT_EXECUTE>`,
    },
  ];
}

export function loadTemplateSectionSystemPrompt(
  promptPath: string | URL = DEFAULT_TEMPLATE_SECTION_SYSTEM_PROMPT_PATH,
): string {
  const prompt = readFileSync(promptPath, "utf8").trim();
  if (!prompt) throw new Error(`Template section system prompt is empty: ${String(promptPath)}`);
  return prompt;
}

function renderTemplateSectionSystemPrompt(
  prompt: string,
  input: TemplateSectionPromptInput,
): string {
  const values: Record<string, string> = {
    DOCUMENT_ID: input.documentId,
    BASE_REVISION: String(input.baseRevision),
    OPERATION_TYPE: input.templateSection.operationType,
    TARGET_ID: input.templateSection.targetId,
  };
  let rendered = prompt;
  for (const [name, value] of Object.entries(values)) {
    const token = `{{${name}}}`;
    if (!rendered.includes(token))
      throw new Error(`Template section prompt is missing token ${token}`);
    rendered = rendered.replaceAll(token, value);
  }
  const unknownToken = rendered.match(/\{\{[A-Z0-9_]+\}\}/u)?.[0];
  if (unknownToken)
    throw new Error(`Template section prompt contains unknown token ${unknownToken}`);
  return rendered;
}

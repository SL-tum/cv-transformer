import type { DocumentRoot, RichDocument } from "../model/core/document.js";
import {
  previewOperations,
  TEMPLATE_OPERATION_SCHEMA_VERSION,
  type LlmGenerationResult,
  type LlmProvider,
} from "../llm-integration/index.js";
import type {
  LlmTemplateDocument,
  TemplateBindingMap,
  TemplateOperation,
} from "../llm-template/index.js";
import type { CandidateFieldPlan, CandidateGroundTruth, CandidateMappingResult } from "./model.js";
import { extractEditableTemplateSections } from "./template-section-context.js";
import { buildTemplateSectionPrompt } from "./template-section-prompt.js";
import { normalizeLlmTemplateOperation } from "./operation-normalizer.js";
import { evaluateOperationLayoutFit } from "./layout-fit.js";

export interface MapCandidateOptions {
  language?: string;
  requireCompleteCoverage?: boolean;
  maxLayoutRetries?: number;
}

export async function mapCandidateToTemplate(
  provider: LlmProvider,
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  groundTruth: CandidateGroundTruth,
  options: MapCandidateOptions = {},
): Promise<CandidateMappingResult> {
  const templateSections = extractEditableTemplateSections(template);
  const operations: TemplateOperation[] = [];
  const generations: LlmGenerationResult[] = [];

  for (const templateSection of templateSections) {
    let sectionOperations: TemplateOperation[] | undefined;
    let layoutFeedback: { reasons: string[]; previousOperation: TemplateOperation } | undefined;
    const maxAttempts = 1 + (options.maxLayoutRetries ?? 1);
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const messages = buildTemplateSectionPrompt({
        candidateText: groundTruth.fullText,
        templateSection,
        templateOverview: templateSections,
        documentId: template.documentId,
        baseRevision: template.revision,
        previousOperations: operations,
        ...(layoutFeedback ? { layoutFeedback } : {}),
      });
      const generation = await provider.generateOperations({
        template,
        userInput: {
          currentTargetId: templateSection.targetId,
          candidateId: groundTruth.candidateId,
        },
        allowedOperations: [templateSection.operationType],
        context: {
          language: options.language ?? "en",
          ...(templateSection.purpose ? { purpose: templateSection.purpose } : {}),
        },
        messages,
      });
      generations.push(generation);
      const candidateOperations = requireSectionOperations(generation, templateSection).map(
        normalizeLlmTemplateOperation,
      );
      const failedFit = candidateOperations
        .map((candidateOperation) => ({
          candidateOperation,
          fit: evaluateOperationLayoutFit(templateSection, candidateOperation),
        }))
        .find(({ fit }) => !fit.fits);
      if (!failedFit) {
        sectionOperations = candidateOperations;
        break;
      }
      if (attempt === maxAttempts) {
        throw new Error(
          `LLM content does not fit target ${templateSection.targetId} after ${maxAttempts} attempts: ${failedFit.fit.reasons.join(" ")}`,
        );
      }
      layoutFeedback = {
        reasons: failedFit.fit.reasons,
        previousOperation: failedFit.candidateOperation,
      };
    }
    if (!sectionOperations?.length)
      throw new Error(`No operation generated for ${templateSection.targetId}`);
    operations.push(...sectionOperations);
  }

  const envelope = {
    schemaVersion: TEMPLATE_OPERATION_SCHEMA_VERSION,
    documentId: template.documentId,
    baseRevision: template.revision,
    operations,
  };
  const preview = previewOperations(document, template, bindings, envelope, [
    ...new Set(templateSections.map((section) => section.operationType)),
  ]);
  const operationsByTarget = new Map<string, TemplateOperation[]>();
  for (const operation of operations) {
    const targetOperations = operationsByTarget.get(operation.targetId) ?? [];
    targetOperations.push(operation);
    operationsByTarget.set(operation.targetId, targetOperations);
  }
  const fieldPlan: CandidateFieldPlan[] = templateSections.map((section) => {
    const targetOperations = operationsByTarget.get(section.targetId);
    return targetOperations?.length
      ? {
          targetId: section.targetId,
          ...(section.label ? { label: section.label } : {}),
          nodeType:
            section.operationType === "setList"
              ? "list"
              : section.operationType === "appendCollectionItem"
                ? "collection"
                : "text",
          status: "filled",
          operation: targetOperations[0]!,
          operations: targetOperations,
        }
      : {
          targetId: section.targetId,
          ...(section.label ? { label: section.label } : {}),
          nodeType: section.operationType === "setList" ? "list" : "text",
          status: "missing",
        };
  });
  const missingTargetIds = fieldPlan
    .filter((field) => field.status === "missing")
    .map((field) => field.targetId);
  if ((options.requireCompleteCoverage ?? true) && missingTargetIds.length) {
    throw new Error(
      `LLM did not return operations for editable targets: ${missingTargetIds.join(", ")}`,
    );
  }

  const warnings = generations.flatMap((generation) => generation.warnings ?? []);
  const lastMetadata = generations.at(-1)?.metadata;
  return {
    generation: {
      operations,
      ...(warnings.length ? { warnings } : {}),
      ...(lastMetadata ? { metadata: lastMetadata } : {}),
    },
    generations,
    operations,
    fieldPlan,
    missingTargetIds,
    preview,
  };
}

function requireSectionOperations(
  generation: LlmGenerationResult,
  section: ReturnType<typeof extractEditableTemplateSections>[number],
): TemplateOperation[] {
  const expectedCount = section.operationType === "appendCollectionItem" ? "one or more" : "one";
  if (
    generation.operations.length < 1 ||
    (section.operationType !== "appendCollectionItem" && generation.operations.length !== 1)
  ) {
    throw new Error(
      `LLM must return ${expectedCount} operation for ${section.targetId}; received ${generation.operations.length}`,
    );
  }
  for (const operation of generation.operations) {
    if (operation.targetId !== section.targetId) {
      throw new Error(
        `LLM returned target ${operation.targetId}; expected current target ${section.targetId}`,
      );
    }
    if (operation.op !== section.operationType) {
      throw new Error(
        `LLM returned operation ${operation.op}; expected ${section.operationType} for ${section.targetId}`,
      );
    }
  }
  return generation.operations;
}

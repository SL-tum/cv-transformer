import type { RichDocument } from "../model/core/document.js";
import type { DocumentRoot } from "../model/core/document.js";
import { exportDocx } from "../exporters/docx-exporter.js";
import { exportPptx } from "../exporters/pptx-exporter.js";
import { loadOpcPackage, validateOpcPackage } from "../ooxml/opc/package.js";
import {
  evaluateTemplateQuality,
  executeTemplateOperations,
  type ConstraintWarning,
  type LlmTemplateDocument,
  type TemplateBindingMap,
  type TemplateOperation,
  type TemplateQualityReport,
} from "../llm-template/index.js";
import type { TemplateOperationEnvelope } from "./structured-output.js";
import type { LlmGenerationRequest, LlmGenerationResult, LlmProvider } from "./provider.js";
import { TEMPLATE_OPERATION_SCHEMA_VERSION } from "./structured-output.js";

export interface MutationEstimate {
  textReplacements: number;
  nodesInserted: number;
  nodesRemoved: number;
  resourcesAdded: number;
}
export interface OperationPreview {
  valid: boolean;
  operations: number;
  affectedTemplateNodes: string[];
  warnings: ConstraintWarning[];
  estimatedMutations: MutationEstimate;
  errors: string[];
}
export interface AtomicExecutionResult {
  preview: OperationPreview;
  quality: TemplateQualityReport;
}

export class LlmOperationExecutor {
  constructor(
    private readonly document: RichDocument<DocumentRoot>,
    private readonly template: LlmTemplateDocument,
    private readonly bindings: TemplateBindingMap,
    private readonly allowedOperations?: string[],
  ) {}
  preview(envelope: TemplateOperationEnvelope): OperationPreview {
    return previewOperations(
      this.document,
      this.template,
      this.bindings,
      envelope,
      this.allowedOperations,
    );
  }
  execute(envelope: TemplateOperationEnvelope): AtomicExecutionResult {
    return executeOperationsAtomically(
      this.document,
      this.template,
      this.bindings,
      envelope,
      this.allowedOperations,
    );
  }
}

export async function generateOperationPreview(
  provider: LlmProvider,
  request: LlmGenerationRequest,
  document: RichDocument<DocumentRoot>,
  bindings: TemplateBindingMap,
): Promise<{
  generation: LlmGenerationResult;
  envelope: TemplateOperationEnvelope;
  preview: OperationPreview;
}> {
  const generation = await provider.generateOperations(request);
  const envelope: TemplateOperationEnvelope = {
    schemaVersion: TEMPLATE_OPERATION_SCHEMA_VERSION,
    documentId: request.template.documentId,
    baseRevision: request.template.revision,
    operations: generation.operations,
  };
  const preview = previewOperations(
    document,
    request.template,
    bindings,
    envelope,
    request.allowedOperations,
  );
  return { generation, envelope, preview };
}

export function previewOperations(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  envelope: TemplateOperationEnvelope,
  allowedOperations?: string[],
): OperationPreview {
  const errors = envelopeErrors(template, envelope, allowedOperations);
  const estimate = estimateMutations(envelope.operations);
  if (errors.length) return preview(false, envelope.operations, [], estimate, errors);
  const documentClone = structuredClone(document);
  const templateClone = structuredClone(template);
  const bindingsClone = structuredClone(bindings);
  try {
    const result = executeTemplateOperations(
      documentClone,
      templateClone,
      bindingsClone,
      envelope.operations,
    );
    const quality = evaluateTemplateQuality(templateClone, bindingsClone, documentClone);
    if (!quality.passed) errors.push(...quality.issues);
    validateExport(documentClone);
    return preview(errors.length === 0, envelope.operations, result.warnings, estimate, errors);
  } catch (error) {
    errors.push(error instanceof Error ? error.message : String(error));
    return preview(false, envelope.operations, [], estimate, errors);
  }
}

export function executeOperationsAtomically(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  envelope: TemplateOperationEnvelope,
  allowedOperations?: string[],
): AtomicExecutionResult {
  const operationPreview = previewOperations(
    document,
    template,
    bindings,
    envelope,
    allowedOperations,
  );
  if (!operationPreview.valid)
    throw new Error(`Operation preview failed: ${operationPreview.errors.join("; ")}`);
  const documentSnapshot = structuredClone(document);
  const templateSnapshot = structuredClone(template);
  const bindingsSnapshot = structuredClone(bindings);
  try {
    executeTemplateOperations(document, template, bindings, envelope.operations);
    const quality = evaluateTemplateQuality(template, bindings, document);
    if (!quality.passed) throw new Error(`Quality gates failed: ${quality.issues.join("; ")}`);
    validateExport(structuredClone(document));
    return { preview: operationPreview, quality };
  } catch (error) {
    restore(document, documentSnapshot);
    restore(template, templateSnapshot);
    restore(bindings, bindingsSnapshot);
    throw error;
  }
}

function envelopeErrors(
  template: LlmTemplateDocument,
  envelope: TemplateOperationEnvelope,
  allowed?: string[],
): string[] {
  const errors: string[] = [];
  if (envelope.documentId !== template.documentId)
    errors.push(`documentId mismatch: expected ${template.documentId}`);
  if (envelope.baseRevision !== template.revision)
    errors.push(
      `baseRevision mismatch: expected ${template.revision}, received ${envelope.baseRevision}`,
    );
  if (allowed)
    for (const operation of envelope.operations)
      if (!allowed.includes(operation.op))
        errors.push(`Operation is not globally allowed: ${operation.op}`);
  return errors;
}
function estimateMutations(operations: TemplateOperation[]): MutationEstimate {
  const result: MutationEstimate = {
    textReplacements: 0,
    nodesInserted: 0,
    nodesRemoved: 0,
    resourcesAdded: 0,
  };
  for (const operation of operations) {
    if (operation.op === "setText") result.textReplacements++;
    else if (operation.op === "setList") result.textReplacements += operation.items.length;
    else if (operation.op === "appendListItem") result.nodesInserted++;
    else if (operation.op === "removeListItem" || operation.op === "removeCollectionItem")
      result.nodesRemoved++;
    else if (operation.op === "appendCollectionItem")
      result.nodesInserted += 1 + Object.keys(operation.value).length;
    else if (operation.op === "updateCollectionItem")
      result.textReplacements += Object.keys(operation.value).length;
  }
  return result;
}
function preview(
  valid: boolean,
  operations: TemplateOperation[],
  warnings: ConstraintWarning[],
  estimatedMutations: MutationEstimate,
  errors: string[],
): OperationPreview {
  return {
    valid,
    operations: operations.length,
    affectedTemplateNodes: [...new Set(operations.map((item) => item.targetId))],
    warnings,
    estimatedMutations,
    errors,
  };
}
function validateExport(document: RichDocument<DocumentRoot>) {
  const bytes =
    document.format === "docx"
      ? exportDocx(document as Parameters<typeof exportDocx>[0])
      : exportPptx(document as Parameters<typeof exportPptx>[0]);
  const issues = validateOpcPackage(loadOpcPackage(bytes));
  if (issues.length) throw new Error(`Exported OPC validation failed: ${issues.join("; ")}`);
}
function restore<T extends object>(target: T, snapshot: T) {
  for (const key of Object.keys(target) as Array<keyof T>) delete target[key];
  Object.assign(target, snapshot);
}

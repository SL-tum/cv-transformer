import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import { validateDocument } from "../../validation.js";
import { nativeStoreToPackage, validateOpcPackage } from "../../ooxml/opc/package.js";
import type { TemplateBindingMap } from "../bindings/binding-map.js";
import { flatten, validateBindingMap } from "../bindings/binding-validator.js";
import type { LlmTemplateDocument } from "../model/template-document.js";

export interface TemplateQualityReport {
  passed: boolean;
  editableNodes: number;
  boundEditableNodes: number;
  editableBindingCoverage: number;
  explicitMarkerCoverage: number;
  inferredBindingCoverage: number;
  averageInferenceConfidence: number;
  needsReview: number;
  duplicateTemplateIds: number;
  unresolvedBindings: number;
  collectionsWithoutPrototype: number;
  brokenRdtReferences: number;
  brokenOpcRelationships: number;
  issues: string[];
}
export function evaluateTemplateQuality(
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  document: RichDocument<DocumentRoot>,
): TemplateQualityReport {
  const nodes = flatten(template.root);
  const editable = nodes.filter((node) => node.editable);
  const bound = editable.filter((node) => Boolean(bindings.bindings[node.id]));
  const sourced = nodes.filter((node) => extractionSource(node));
  const explicit = sourced.filter((node) => extractionSource(node) === "explicitMarker");
  const inferred = sourced.filter((node) => extractionSource(node) !== "explicitMarker");
  const confidences = inferred
    .map((node) => extractionConfidence(node))
    .filter((value): value is number => value !== undefined);
  const needsReview = nodes.filter((node) => extractionStatus(node) === "needsReview").length;
  const bindingIssues = validateBindingMap(template, bindings, document);
  const rdtIssues = validateDocument(document);
  const opcIssues = document.nativeStore
    ? validateOpcPackage(nativeStoreToPackage(document.nativeStore))
    : ["NativeStore is missing"];
  const duplicateTemplateIds = bindingIssues.filter(
    (item) => item.code === "duplicate-template-id",
  ).length;
  const unresolvedBindings = bindingIssues.filter(
    (item) => item.code === "unbound-editable-node" || item.code === "unresolved-rdt-node",
  ).length;
  const collectionsWithoutPrototype = bindingIssues.filter(
    (item) => item.code === "missing-prototype",
  ).length;
  const brokenRdtReferences = rdtIssues.filter((item) => item.severity === "error").length;
  const issues = [
    ...bindingIssues.map((item) => item.message),
    ...rdtIssues.map((item) => item.message),
    ...opcIssues,
  ];
  const editableBindingCoverage = editable.length ? bound.length / editable.length : 1;
  return {
    passed:
      editableBindingCoverage === 1 &&
      duplicateTemplateIds === 0 &&
      unresolvedBindings === 0 &&
      collectionsWithoutPrototype === 0 &&
      brokenRdtReferences === 0 &&
      opcIssues.length === 0,
    editableNodes: editable.length,
    boundEditableNodes: bound.length,
    editableBindingCoverage,
    explicitMarkerCoverage: sourced.length ? explicit.length / sourced.length : 0,
    inferredBindingCoverage: sourced.length ? inferred.length / sourced.length : 0,
    averageInferenceConfidence: confidences.length
      ? confidences.reduce((sum, value) => sum + value, 0) / confidences.length
      : 1,
    needsReview,
    duplicateTemplateIds,
    unresolvedBindings,
    collectionsWithoutPrototype,
    brokenRdtReferences,
    brokenOpcRelationships: opcIssues.length,
    issues,
  };
}
function extraction(node: { metadata?: Record<string, unknown> }) {
  const value = node.metadata?.extraction;
  return value && typeof value === "object" ? (value as Record<string, unknown>) : undefined;
}
function extractionSource(node: { metadata?: Record<string, unknown> }) {
  return typeof node.metadata?.extractionSource === "string"
    ? node.metadata.extractionSource
    : typeof extraction(node)?.source === "string"
      ? (extraction(node)!.source as string)
      : undefined;
}
function extractionConfidence(node: { metadata?: Record<string, unknown> }) {
  const value = extraction(node)?.confidence;
  return typeof value === "number" ? value : undefined;
}
function extractionStatus(node: { metadata?: Record<string, unknown> }) {
  const value = extraction(node)?.status;
  return typeof value === "string" ? value : undefined;
}

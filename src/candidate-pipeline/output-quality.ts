import { importDocx } from "../importers/docx-importer.js";
import { importPptx } from "../importers/pptx-importer.js";
import type { RichDocument } from "../model/core/document.js";
import type { DocumentRoot } from "../model/core/document.js";
import type { TemplateOperation } from "../llm-template/index.js";
import { loadOpcPackage, validateOpcPackage } from "../ooxml/opc/package.js";
import { validateDocument } from "../validation.js";

export interface OutputQualityReport {
  valid: boolean;
  errors: string[];
  warnings: string[];
  checks: {
    opcValid: boolean;
    rdtValid: boolean;
    operationContentPresent: boolean;
    whitespaceValid: boolean;
    blankTableRows: number;
  };
}

export function assessGeneratedOutput(
  bytes: Uint8Array,
  format: "docx" | "pptx",
  operations: TemplateOperation[],
): OutputQualityReport {
  const errors: string[] = [];
  const warnings: string[] = [];
  if (!bytes.length) errors.push("Exported document is empty.");

  let opcValid = false;
  try {
    const opcIssues = validateOpcPackage(loadOpcPackage(bytes));
    opcValid = opcIssues.length === 0;
    errors.push(...opcIssues.map((issue) => `OPC: ${issue}`));
  } catch (error) {
    errors.push(`OPC could not be loaded: ${message(error)}`);
  }

  let document: RichDocument<DocumentRoot> | undefined;
  let rdtValid = false;
  try {
    document = format === "docx" ? importDocx(bytes) : importPptx(bytes);
    const rdtErrors = validateDocument(document).filter((issue) => issue.severity === "error");
    rdtValid = rdtErrors.length === 0;
    errors.push(...rdtErrors.map((issue) => `RDT ${issue.code}: ${issue.message}`));
  } catch (error) {
    errors.push(`Exported ${format.toUpperCase()} could not be re-imported: ${message(error)}`);
  }

  const exportedTextValues = document ? collectDocumentTextValues(document) : [];
  const exportedText = exportedTextValues.join("\n");
  const missingValues = operationTextValues(operations).filter(
    (value) => value && !exportedText.includes(value),
  );
  const operationContentPresent = missingValues.length === 0;
  if (missingValues.length) {
    errors.push(
      `Generated content is missing after export: ${missingValues
        .slice(0, 3)
        .map((value) => JSON.stringify(value.slice(0, 80)))
        .join(", ")}`,
    );
  }

  const whitespaceValid = exportedTextValues.every(
    (value) => !/[\v\f]/u.test(value) && !/\n[ \t]*\n/u.test(value),
  );
  if (!whitespaceValid) errors.push("Output contains control whitespace or repeated blank lines.");

  const blankTableRows = document?.format === "docx" ? countBlankTableRows(document) : 0;
  if (blankTableRows) errors.push(`Output contains ${blankTableRows} completely blank table rows.`);

  return {
    valid: errors.length === 0,
    errors,
    warnings,
    checks: {
      opcValid,
      rdtValid,
      operationContentPresent,
      whitespaceValid,
      blankTableRows,
    },
  };
}

function operationTextValues(operations: TemplateOperation[]): string[] {
  return operations.flatMap((operation) => {
    if (operation.op === "setText" || operation.op === "appendListItem") return [operation.value];
    if (operation.op === "setList") return operation.items;
    if (operation.op === "appendCollectionItem" || operation.op === "updateCollectionItem") {
      return Object.values(operation.value).flatMap((value) =>
        Array.isArray(value) ? value : [value],
      );
    }
    return [];
  });
}

function collectDocumentTextValues(document: RichDocument<DocumentRoot>): string[] {
  const values: string[] = [];
  const visit = (value: unknown): void => {
    if (!value || typeof value !== "object" || value instanceof Uint8Array) return;
    const record = value as Record<string, unknown>;
    if (record.type === "textRun" && typeof record.text === "string") values.push(record.text);
    for (const [key, child] of Object.entries(record)) {
      if (["source", "native", "metadata", "style", "nativeStore", "patchPlan"].includes(key))
        continue;
      if (Array.isArray(child)) child.forEach(visit);
      else visit(child);
    }
  };
  visit(document.root);
  return values;
}

function countBlankTableRows(document: RichDocument<DocumentRoot>): number {
  let count = 0;
  const visit = (value: unknown): void => {
    if (!value || typeof value !== "object" || value instanceof Uint8Array) return;
    const record = value as Record<string, unknown>;
    if (record.type === "tableRow" && Array.isArray(record.cells)) {
      const text = collectText(record.cells).trim();
      if (!text) count++;
    }
    for (const [key, child] of Object.entries(record)) {
      if (["source", "native", "metadata", "style"].includes(key)) continue;
      if (Array.isArray(child)) child.forEach(visit);
      else visit(child);
    }
  };
  visit(document.root);
  return count;
}

function collectText(value: unknown): string {
  const values: string[] = [];
  const visit = (item: unknown): void => {
    if (!item || typeof item !== "object" || item instanceof Uint8Array) return;
    const record = item as Record<string, unknown>;
    if (record.type === "textRun" && typeof record.text === "string") values.push(record.text);
    for (const child of Object.values(record))
      Array.isArray(child) ? child.forEach(visit) : visit(child);
  };
  visit(value);
  return values.join(" ");
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

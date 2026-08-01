import type { TemplateOperation } from "../llm-template/index.js";
import type { TemplateSectionContext } from "./template-section-context.js";

export interface LayoutFitResult {
  fits: boolean;
  estimatedLines: number;
  characters: number;
  items: number;
  reasons: string[];
}

export function evaluateOperationLayoutFit(
  section: TemplateSectionContext,
  operation: TemplateOperation,
): LayoutFitResult {
  const capacity = section.layoutCapacity;
  const values = operationValues(operation);
  const characters = values.reduce((sum, value) => sum + value.length, 0);
  const items = operation.op === "setList" ? operation.items.length : 1;
  if (!capacity) {
    return { fits: true, estimatedLines: 0, characters, items, reasons: [] };
  }
  const estimatedLines = values.reduce(
    (sum, value) =>
      sum +
      Math.max(
        1,
        value
          .split("\n")
          .reduce(
            (lineSum, line) =>
              lineSum + Math.max(1, Math.ceil(line.length / capacity.estimatedCharactersPerLine)),
            0,
          ),
      ),
    0,
  );
  const reasons: string[] = [];
  if (estimatedLines > capacity.estimatedMaxLines) {
    reasons.push(
      `Estimated ${estimatedLines} lines exceeds the available ${capacity.estimatedMaxLines} lines.`,
    );
  }
  if (characters > capacity.recommendedCharacters) {
    reasons.push(
      `${characters} characters exceeds the recommended ${capacity.recommendedCharacters} characters.`,
    );
  }
  if (operation.op === "setList" && items > capacity.estimatedMaxLines) {
    reasons.push(`${items} list items cannot fit in ${capacity.estimatedMaxLines} lines.`);
  }
  return { fits: reasons.length === 0, estimatedLines, characters, items, reasons };
}

function operationValues(operation: TemplateOperation): string[] {
  if (operation.op === "setText") return [operation.value];
  if (operation.op === "setList") return operation.items;
  if (operation.op === "appendListItem") return [operation.value];
  if (operation.op === "appendCollectionItem" || operation.op === "updateCollectionItem") {
    return Object.values(operation.value).flatMap((value) =>
      Array.isArray(value) ? value : [value],
    );
  }
  return [];
}

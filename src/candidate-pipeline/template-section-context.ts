import type { LlmTemplateDocument, TemplateNode } from "../llm-template/index.js";
import type { PptxLayoutCapacity } from "../llm-template/index.js";

export interface TemplateSectionContext {
  targetId: string;
  label?: string;
  existingText?: string;
  nearbyText?: string[];
  parentLabel?: string;
  sectionType?: string;
  purpose?: string;
  layoutCapacity?: PptxLayoutCapacity;
  itemSchema?: Array<{ key: string; label?: string; type: string }>;
  operationType: "setText" | "setList" | "appendCollectionItem";
}

export function extractEditableTemplateSections(
  template: LlmTemplateDocument,
): TemplateSectionContext[] {
  const sections: TemplateSectionContext[] = [];

  const visit = (node: TemplateNode, parent?: TemplateNode): void => {
    if (isSupportedEditableNode(node)) {
      const siblings = parent?.type === "container" ? parent.children : [];
      const siblingIndex = siblings.indexOf(node);
      const nearbyText = siblings
        .filter((_, index) => index !== siblingIndex && Math.abs(index - siblingIndex) <= 2)
        .flatMap(descriptiveText)
        .filter(Boolean);
      const label = node.label ?? node.id;
      const currentText = existingText(node);
      const layoutCapacity = node.metadata?.layoutCapacity as PptxLayoutCapacity | undefined;
      const itemSchema = node.metadata?.itemSchema as TemplateSectionContext["itemSchema"];
      sections.push({
        targetId: node.id,
        ...(node.label ? { label: node.label } : {}),
        ...(currentText ? { existingText: currentText } : {}),
        ...(nearbyText.length ? { nearbyText: [...new Set(nearbyText)] } : {}),
        ...(parent?.label ? { parentLabel: parent.label } : {}),
        sectionType: inferSectionType(label, node.type),
        purpose: inferPurpose(label, node.type),
        ...(layoutCapacity ? { layoutCapacity } : {}),
        ...(itemSchema?.length ? { itemSchema } : {}),
        operationType:
          node.type === "list"
            ? "setList"
            : node.type === "collection"
              ? "appendCollectionItem"
              : "setText",
      });
    }
    if (node.type === "container") {
      node.children.forEach((child) => visit(child, node));
    }
  };

  visit(template.root);
  return sections;
}

function isSupportedEditableNode(node: TemplateNode): boolean {
  return (
    node.editable &&
    (node.type === "text" ||
      node.type === "list" ||
      (node.type === "collection" && Boolean(node.prototypeId)))
  );
}

function existingText(node: TemplateNode): string | undefined {
  if (node.type === "text") return node.value || node.placeholder;
  if (node.type === "list") {
    const value = node.items
      .map((item) => item.value)
      .filter(Boolean)
      .join("\n");
    return value || undefined;
  }
  return undefined;
}

function descriptiveText(node: TemplateNode): string[] {
  const values = [node.label];
  if (node.type === "text") values.push(node.value || node.placeholder);
  if (node.type === "list") values.push(...node.items.map((item) => item.value));
  return values.filter((value): value is string => Boolean(value));
}

function inferSectionType(label: string, nodeType: TemplateNode["type"]): string {
  if (/summary|profile|overview|introduction/iu.test(label)) return "professional-summary";
  if (/skill|capabilit|competenc/iu.test(label)) return "skills";
  if (/experience|employment|assignment|project/iu.test(label)) return "professional-experience";
  if (/client|customer/iu.test(label)) return "clients";
  if (/certificate|certification|qualification/iu.test(label)) return "certifications";
  if (/education|degree|academic/iu.test(label)) return "education";
  if (/language/iu.test(label)) return "languages";
  if (/industr/iu.test(label)) return "industry-experience";
  return nodeType === "list" ? "fact-list" : "text-section";
}

function inferPurpose(label: string, nodeType: TemplateNode["type"]): string {
  if (/summary|profile|overview|introduction/iu.test(label)) {
    return "Summarize the candidate's strongest qualifications and relevant professional value for the target CV.";
  }
  if (/skill|capabilit|competenc/iu.test(label)) {
    return "Select the candidate capabilities that best fit this template skill category.";
  }
  if (/experience|employment|assignment|project/iu.test(label)) {
    return "Present the candidate experience most relevant to this target section, preserving employers, roles and dates.";
  }
  if (/client|customer/iu.test(label)) {
    return "List only explicitly named clients or customers supported by the candidate source facts.";
  }
  if (/certificate|certification|qualification/iu.test(label)) {
    return "List the candidate's explicitly stated professional certifications and qualifications.";
  }
  if (/education|degree|academic/iu.test(label)) {
    return "Present the candidate's relevant education in the structure expected by this template section.";
  }
  if (/language/iu.test(label)) {
    return "List explicitly stated languages and proficiency levels.";
  }
  if (/industr/iu.test(label)) {
    return "Select industries explicitly supported by the candidate's work and project facts.";
  }
  return nodeType === "list"
    ? "Select and organize source facts that belong in this target list."
    : "Produce final template-ready content for this target section using supported source facts.";
}

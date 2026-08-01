import type { RichDocument } from "../../model/core/document.js";
import type { Transform } from "../../model/core/primitives.js";
import type { GraphicNode } from "../../model/drawing/index.js";
import type { PresentationDocumentRoot, SlideElementNode } from "../../model/presentation/index.js";
import type { ParagraphNode, TextRunNode } from "../../model/text/index.js";
import type { FlowBlockNode, FlowDocumentRoot } from "../../model/word/index.js";
import type { TemplateCandidate } from "./candidate.js";

export function discoverPptxCandidates(
  document: RichDocument<PresentationDocumentRoot>,
): TemplateCandidate[] {
  const result: TemplateCandidate[] = [];
  document.root.slides.forEach((slide, slideIndex) =>
    discoverPptxElements(slide.shapes, slideIndex, result),
  );
  linkNeighbors(result);
  return result;
}
function discoverPptxElements(
  elements: Array<SlideElementNode | GraphicNode>,
  slideIndex: number,
  output: TemplateCandidate[],
) {
  for (const element of elements) {
    if (element.type === "group") {
      const directTitles = element.elements.flatMap((child) => directText(child));
      if (
        directTitles.length === 1 &&
        directTitles[0]!.text.length <= 64 &&
        element.elements.length >= 2
      ) {
        const title = directTitles[0]!;
        const layout = layoutOf(element.transform);
        output.push({
          id: `candidate:group:${element.id}`,
          sourceNodeIds: [element.id, title.paragraph.id, ...title.runs.map((run) => run.id)],
          text: title.text,
          styleFeatures: styleOf(title.paragraph),
          layoutFeatures: { ...layout, slideIndex, containerNodeId: element.id, order: -1 },
          nativeRole: "shapeGroupTitle",
          ...(element.source
            ? { partUri: element.source.partUri, xmlPath: element.source.xmlPath }
            : {}),
          signals: {
            text: { value: title.text, short: true, placeholder: false },
            style: styleOf(title.paragraph),
            geometry: { ...layout, slideIndex, containerNodeId: element.id, order: -1 },
            nativeRole: { role: "shapeGroupTitle" },
          },
        });
      }
      discoverPptxElements(element.elements, slideIndex, output);
      continue;
    }
    if (
      (element.type !== "shape" && element.type !== "placeholder") ||
      !("textBody" in element) ||
      !element.textBody
    )
      continue;
    element.textBody.paragraphs.forEach((paragraph, order) => {
      const textRuns = paragraph.runs.filter((run): run is TextRunNode => run.type === "textRun");
      const placeholderRuns = textRuns.filter((run) => isPlaceholderText(run.text));
      const selectedRuns = placeholderRuns.length ? placeholderRuns : textRuns;
      const text = selectedRuns
        .map((run) => run.text)
        .join("")
        .trim();
      if (!text) return;
      const style = styleOf(paragraph);
      const layout = layoutOf(element.transform);
      const placeholder = isPlaceholderText(text);
      const source = selectedRuns.length === 1 ? selectedRuns[0]!.source : paragraph.source;
      output.push({
        id: placeholderRuns.length
          ? `candidate:${paragraph.id}:placeholder`
          : `candidate:${paragraph.id}`,
        sourceNodeIds: placeholderRuns.length
          ? selectedRuns.map((run) => run.id)
          : [paragraph.id, ...selectedRuns.map((run) => run.id)],
        text,
        styleFeatures: style,
        layoutFeatures: { ...layout, slideIndex, containerNodeId: element.id, order },
        nativeRole: element.type === "placeholder" ? element.placeholderType : element.type,
        ...(source ? { partUri: source.partUri, xmlPath: source.xmlPath } : {}),
        signals: {
          text: { value: text, short: text.length <= 48, placeholder },
          style,
          geometry: { ...layout, slideIndex, containerNodeId: element.id, order },
          nativeRole: {
            role: element.type === "placeholder" ? element.placeholderType : element.type,
          },
        },
      });
    });
  }
}

export function discoverDocxCandidates(
  document: RichDocument<FlowDocumentRoot>,
): TemplateCandidate[] {
  const result: TemplateCandidate[] = [];
  document.root.sections.forEach((section, sectionIndex) =>
    discoverDocxBlocks(section.blocks, sectionIndex, result),
  );
  linkNeighbors(result);
  return result;
}
function discoverDocxBlocks(
  blocks: FlowBlockNode[],
  sectionIndex: number,
  output: TemplateCandidate[],
) {
  blocks.forEach((block, order) => {
    if (block.type === "paragraph" && "runs" in block) {
      const textRuns = block.runs.filter((run): run is TextRunNode => run.type === "textRun");
      const text = textRuns
        .map((run) => run.text)
        .join("")
        .trim();
      if (!text) return;
      const style = styleOf(block);
      output.push({
        id: `candidate:${block.id}`,
        sourceNodeIds: [block.id, ...textRuns.map((run) => run.id)],
        text,
        styleFeatures: style,
        layoutFeatures: { containerNodeId: `section:${sectionIndex}`, order },
        nativeRole: block.properties.styleId?.toLowerCase().startsWith("heading")
          ? "heading"
          : block.properties.list
            ? "listParagraph"
            : "paragraph",
        ...(block.source ? { partUri: block.source.partUri, xmlPath: block.source.xmlPath } : {}),
        signals: {
          text: { value: text, short: text.length <= 80, placeholder: isPlaceholderText(text) },
          style,
          geometry: { containerNodeId: `section:${sectionIndex}`, order },
          nativeRole: {
            role: block.properties.styleId?.toLowerCase().startsWith("heading")
              ? "heading"
              : "paragraph",
          },
        },
      });
    } else if (
      (block.type === "structuredDocumentTag" || block.type === "customXml") &&
      "blocks" in block
    )
      discoverDocxBlocks(block.blocks, sectionIndex, output);
  });
}

function styleOf(paragraph: ParagraphNode) {
  const run = paragraph.runs.find((item): item is TextRunNode => item.type === "textRun");
  const size = run?.properties.fontSize?.unit === "pt" ? run.properties.fontSize.value : undefined;
  return {
    ...(run?.properties.bold !== undefined ? { bold: run.properties.bold } : {}),
    ...(size !== undefined ? { fontSizePt: size } : {}),
    ...(paragraph.properties.styleId ? { styleId: paragraph.properties.styleId } : {}),
    emphasized: Boolean(
      run?.properties.bold || paragraph.properties.styleId?.toLowerCase().startsWith("heading"),
    ),
  };
}
function directText(
  element: GraphicNode,
): Array<{ text: string; paragraph: ParagraphNode; runs: TextRunNode[] }> {
  if (element.type !== "shape" || !element.textBody) return [];
  return element.textBody.paragraphs.flatMap((paragraph) => {
    const runs = paragraph.runs.filter((run): run is TextRunNode => run.type === "textRun");
    const text = runs
      .map((run) => run.text)
      .join("")
      .trim();
    return text ? [{ text, paragraph, runs }] : [];
  });
}
function layoutOf(transform?: Transform) {
  return {
    ...(transform?.x ? { x: transform.x.value } : {}),
    ...(transform?.y ? { y: transform.y.value } : {}),
    ...(transform?.width ? { width: transform.width.value } : {}),
    ...(transform?.height ? { height: transform.height.value } : {}),
  };
}
function isPlaceholderText(text: string) {
  return /^\s*(?:\[…\]|\[\.{3}\]|<.+>|lorem ipsum|write .+ here|add .+ here)\s*$/i.test(text);
}
function linkNeighbors(candidates: TemplateCandidate[]) {
  const groups = groupBy(
    candidates,
    (candidate) =>
      `${candidate.layoutFeatures?.slideIndex ?? "doc"}:${candidate.layoutFeatures?.containerNodeId ?? "root"}`,
  );
  for (const group of groups.values()) {
    group.sort((a, b) => (a.layoutFeatures?.order ?? 0) - (b.layoutFeatures?.order ?? 0));
    group.forEach((candidate, index) => {
      candidate.signals.neighborhood = {
        ...(group[index - 1] ? { previousCandidateId: group[index - 1]!.id } : {}),
        ...(group[index + 1] ? { nextCandidateId: group[index + 1]!.id } : {}),
      };
    });
  }
}
function groupBy<T>(values: T[], key: (value: T) => string) {
  const result = new Map<string, T[]>();
  for (const value of values) {
    const id = key(value);
    const items = result.get(id) ?? [];
    items.push(value);
    result.set(id, items);
  }
  return result;
}

import type { RichDocument } from "../../model/core/document.js";
import type { GraphicNode } from "../../model/drawing/index.js";
import type {
  PlaceholderNode,
  PresentationDocumentRoot,
  SlideElementNode,
  SlideLayoutNode,
  SlideMasterNode,
} from "../../model/presentation/index.js";
import type { ParagraphNode, TextRunNode } from "../../model/text/index.js";
import { flatten } from "../bindings/binding-validator.js";
import type { ExtractionResult } from "./extractor.js";

export interface PptxLayoutCapacity {
  widthPt: number;
  heightPt: number;
  fontSizePt: number;
  estimatedCharactersPerLine: number;
  estimatedMaxLines: number;
  recommendedCharacters: number;
  recommendedItems: number;
  sharedShapeTargets: number;
}

export function applyPptxLayoutCapacities(
  document: RichDocument<PresentationDocumentRoot>,
  result: ExtractionResult,
): ExtractionResult {
  const shapeByDescendantId = indexTextShapes(document);
  const targetShapes = new Map<string, TextShapeInfo>();
  for (const node of flatten(result.template.root)) {
    if (!node.editable) continue;
    const binding = result.bindings.bindings[node.id];
    const shape = binding?.sourceNodeIds.map((id) => shapeByDescendantId.get(id)).find(Boolean);
    if (shape && shape.widthEmu > 0 && shape.heightEmu > 0) targetShapes.set(node.id, shape);
  }
  const sharedCounts = new Map<string, number>();
  for (const shape of targetShapes.values()) {
    sharedCounts.set(shape.id, (sharedCounts.get(shape.id) ?? 0) + 1);
  }
  for (const node of flatten(result.template.root)) {
    const shape = targetShapes.get(node.id);
    if (!shape) continue;
    const sharedShapeTargets = sharedCounts.get(shape.id) ?? 1;
    const capacity = estimateCapacity(shape, sharedShapeTargets);
    node.constraints = {
      ...node.constraints,
      soft: {
        recommendedCharacters:
          node.constraints?.soft?.recommendedCharacters ?? capacity.recommendedCharacters,
        recommendedItems: node.constraints?.soft?.recommendedItems ?? capacity.recommendedItems,
        maxLines: node.constraints?.soft?.maxLines ?? capacity.estimatedMaxLines,
        overflowRisk: node.constraints?.soft?.overflowRisk ?? "medium",
        ...(node.constraints?.soft?.recommendedWords !== undefined
          ? { recommendedWords: node.constraints.soft.recommendedWords }
          : {}),
        ...(node.constraints?.soft?.recommendedWordsPerItem !== undefined
          ? { recommendedWordsPerItem: node.constraints.soft.recommendedWordsPerItem }
          : {}),
      },
    };
    node.metadata = { ...node.metadata, layoutCapacity: capacity };
  }
  return result;
}

interface TextShapeInfo {
  id: string;
  widthEmu: number;
  heightEmu: number;
  paragraphs: ParagraphNode[];
}

function indexTextShapes(
  document: RichDocument<PresentationDocumentRoot>,
): Map<string, TextShapeInfo> {
  const result = new Map<string, TextShapeInfo>();
  const visit = (
    elements: Array<SlideElementNode | GraphicNode>,
    layout?: SlideLayoutNode,
    master?: SlideMasterNode,
  ): void => {
    for (const element of elements) {
      if (element.type === "group") {
        visit(element.elements, layout, master);
        continue;
      }
      if ((element.type !== "shape" && element.type !== "placeholder") || !element.textBody)
        continue;
      const inherited =
        element.type === "placeholder"
          ? inheritedPlaceholderGeometry(element, layout, master)
          : undefined;
      const info: TextShapeInfo = {
        id: element.id,
        widthEmu: element.transform?.width?.value ?? inherited?.transform?.width?.value ?? 0,
        heightEmu: element.transform?.height?.value ?? inherited?.transform?.height?.value ?? 0,
        paragraphs: element.textBody.paragraphs,
      };
      result.set(element.id, info);
      for (const paragraph of element.textBody.paragraphs) {
        result.set(paragraph.id, info);
        for (const run of paragraph.runs) result.set(run.id, info);
      }
    }
  };
  for (const slide of document.root.slides) {
    const layout = document.root.slideLayouts.find((item) => item.id === slide.layoutRef);
    const masterRef = slide.masterRef ?? layout?.masterRef;
    const master = document.root.slideMasters.find((item) => item.id === masterRef);
    visit(slide.shapes, layout, master);
  }
  return result;
}

function inheritedPlaceholderGeometry(
  placeholder: PlaceholderNode,
  layout?: SlideLayoutNode,
  master?: SlideMasterNode,
): PlaceholderNode | undefined {
  const layoutPlaceholder = findPlaceholder(layout?.shapes, placeholder);
  const masterPlaceholder = findPlaceholder(master?.shapes, layoutPlaceholder ?? placeholder);
  if (!layoutPlaceholder) return masterPlaceholder;
  if (layoutPlaceholder.transform?.width && layoutPlaceholder.transform.height)
    return layoutPlaceholder;
  return masterPlaceholder ?? layoutPlaceholder;
}

function findPlaceholder(
  elements: SlideElementNode[] | undefined,
  reference: PlaceholderNode,
): PlaceholderNode | undefined {
  const placeholders = elements?.filter(
    (element): element is PlaceholderNode => element.type === "placeholder",
  );
  if (reference.placeholderIndex !== undefined) {
    const indexed = placeholders?.find(
      (element) => element.placeholderIndex === reference.placeholderIndex,
    );
    if (indexed) return indexed;
  }
  return placeholders?.find((element) => element.placeholderType === reference.placeholderType);
}

function estimateCapacity(shape: TextShapeInfo, sharedShapeTargets: number): PptxLayoutCapacity {
  const widthPt = Math.max(24, shape.widthEmu / 12700 - 12);
  const fullHeightPt = Math.max(12, shape.heightEmu / 12700 - 8);
  const heightPt = fullHeightPt / sharedShapeTargets;
  const fontSizePt = fontSize(shape.paragraphs);
  const estimatedCharactersPerLine = Math.max(8, Math.floor(widthPt / (fontSizePt * 0.52)));
  const estimatedMaxLines = Math.max(1, Math.floor(heightPt / (fontSizePt * 1.25)));
  const recommendedCharacters = Math.max(
    12,
    Math.floor(estimatedCharactersPerLine * estimatedMaxLines * 0.85),
  );
  const recommendedItems = Math.max(1, Math.floor(estimatedMaxLines / 2));
  return {
    widthPt: round(widthPt),
    heightPt: round(heightPt),
    fontSizePt: round(fontSizePt),
    estimatedCharactersPerLine,
    estimatedMaxLines,
    recommendedCharacters,
    recommendedItems,
    sharedShapeTargets,
  };
}

function fontSize(paragraphs: ParagraphNode[]): number {
  const sizes = paragraphs
    .flatMap((paragraph) => paragraph.runs)
    .filter((run): run is TextRunNode => run.type === "textRun")
    .map((run) => run.properties.fontSize)
    .filter((size) => size?.unit === "pt")
    .map((size) => size!.value);
  return sizes.length ? Math.min(...sizes) : 10;
}

function round(value: number): number {
  return Math.round(value * 10) / 10;
}

import type { TemplateBinding } from "../bindings/binding.js";
import type { ContainerNode, ListNode, TemplateNode, TextNode } from "../model/template-node.js";
import type { TemplateMarker } from "../markers/marker.js";
import type {
  CandidateExtractionResult,
  ExtractionDecision,
  TemplateCandidate,
} from "./candidate.js";
import type { ExtractionOptions, ExtractionResult } from "./extractor.js";

export function mergeHybridExtraction(
  base: ExtractionResult,
  allCandidates: TemplateCandidate[],
  decisions: ExtractionDecision[],
  markers: TemplateMarker[],
  options: ExtractionOptions,
): ExtractionResult {
  const protectedPaths = markers.map((marker) => `${marker.partUri}\u001f${marker.xmlPath}`);
  const candidates = allCandidates.filter(
    (candidate) =>
      !candidate.partUri ||
      !candidate.xmlPath ||
      !protectedPaths.some((path) =>
        `${candidate.partUri}\u001f${candidate.xmlPath}`.startsWith(path),
      ),
  );
  const available = new Map(candidates.map((candidate) => [candidate.id, candidate]));
  const acceptedDecisions = decisions.filter(
    (decision) =>
      decision.status !== "rejected" &&
      decision.proposal.valueCandidateIds?.every((id) => available.has(id)),
  );
  const existingIds = new Set(base.template.root.children.map((node) => node.id));
  const included: ExtractionDecision[] = [];
  for (const decision of acceptedDecisions) {
    const proposal = decision.proposal;
    let id = proposal.id;
    let sequence = 2;
    while (existingIds.has(id)) id = `${proposal.id}-${sequence++}`;
    existingIds.add(id);
    const valueCandidates = (proposal.valueCandidateIds ?? proposal.candidateIds)
      .map((candidateId) => available.get(candidateId))
      .filter((candidate): candidate is TemplateCandidate => Boolean(candidate));
    if (!valueCandidates.length) continue;
    const value = valueCandidates.map((candidate) => candidate.text ?? "").join("\n");
    const editable = decision.status === "accepted";
    const extraction = {
      source: proposal.source,
      confidence: proposal.confidence,
      status: decision.status,
      evidence: proposal.evidence.map((item) => item.description),
      candidateIds: proposal.candidateIds,
    };
    let node: TemplateNode;
    if (proposal.type === "container")
      node = {
        id,
        type: "container",
        label: proposal.label ?? id,
        editable: false,
        children: [],
        metadata: { extraction },
      } satisfies ContainerNode;
    else if (proposal.type === "list") {
      const items = proposal.emptySection
        ? []
        : valueCandidates
            .flatMap((candidate) =>
              candidate.signals.text?.placeholder
                ? []
                : (candidate.text ?? "").split(/\r?\n/).filter(Boolean),
            )
            .map((item, index) => ({ id: `${id}:item:${index + 1}`, value: item }));
      node = {
        id,
        type: "list",
        label: proposal.label ?? id,
        editable,
        repeatable: true,
        items,
        metadata: { extraction },
      } satisfies ListNode;
    } else {
      const placeholder = valueCandidates.find(
        (candidate) => candidate.signals.text?.placeholder,
      )?.text;
      node = {
        id,
        type: "text",
        label: proposal.label ?? id,
        editable,
        value:
          proposal.emptySection ||
          valueCandidates.every((candidate) => candidate.signals.text?.placeholder)
            ? ""
            : value,
        ...(placeholder && !proposal.emptySection ? { placeholder } : {}),
        metadata: { extraction },
      } satisfies TextNode;
    }
    const parent = proposal.parentId
      ? findContainer(base.template.root, proposal.parentId)
      : undefined;
    if (parent) parent.children.push(node);
    else base.template.root.children.push(node);
    base.bindings.bindings[id] =
      proposal.emptySection && proposal.insertionLayout
        ? insertionBinding(id, valueCandidates[0]!, proposal.insertionLayout)
        : bindingForCandidates(
            id,
            valueCandidates,
            base.template.sourceFormat,
            proposal.type === "list",
            valueCandidates.every((candidate) => Boolean(candidate.signals.text?.placeholder)),
          );
    if (decision.status === "needsReview")
      base.warnings.push(
        `Inferred field ${id} needs review (confidence ${proposal.confidence.toFixed(2)})`,
      );
    included.push({ ...decision, proposal: { ...proposal, id } });
  }
  const proposals = decisions.map((decision) => decision.proposal);
  const extraction: CandidateExtractionResult = {
    candidates,
    proposals,
    decisions: decisions.map(
      (decision) =>
        included.find(
          (item) => item.proposal.candidateIds.join() === decision.proposal.candidateIds.join(),
        ) ?? decision,
    ),
  };
  return { ...base, extraction };
}
function bindingForCandidates(
  id: string,
  candidates: TemplateCandidate[],
  format: "docx" | "pptx",
  list: boolean,
  allowEmpty = false,
): TemplateBinding {
  return {
    templateNodeId: id,
    sourceNodeIds: [...new Set(candidates.flatMap((candidate) => candidate.sourceNodeIds))],
    sourceFormat: format,
    representation: list ? "paragraphList" : format === "pptx" ? "shapeText" : "joinedRuns",
    readStrategy: list ? { type: "collectParagraphs" } : { type: "collectText" },
    writeStrategy: list
      ? { type: "replaceParagraphList", ...(allowEmpty ? { allowEmpty: true } : {}) }
      : format === "pptx"
        ? { type: "replaceShapeText" }
        : { type: "replaceTextPreservingRuns" },
    formatStrategy: { type: "preserveExisting" },
    locations: candidates.map((candidate) => ({
      nodeId: candidate.sourceNodeIds[0] ?? candidate.id,
      ...(candidate.partUri ? { partUri: candidate.partUri } : {}),
      ...(candidate.xmlPath ? { xmlPath: candidate.xmlPath } : {}),
      role: "paragraph" as const,
    })),
  };
}
function insertionBinding(
  id: string,
  anchor: TemplateCandidate,
  layout: { x: number; y: number; width: number; height: number },
): TemplateBinding {
  return {
    templateNodeId: id,
    sourceNodeIds: [anchor.sourceNodeIds[0] ?? anchor.id],
    sourceFormat: "pptx",
    representation: "shapeText",
    readStrategy: { type: "collectText" },
    writeStrategy: { type: "insertShapeText", ...layout },
    formatStrategy: { type: "preserveExisting" },
    locations: [
      {
        nodeId: anchor.sourceNodeIds[0] ?? anchor.id,
        ...(anchor.partUri ? { partUri: anchor.partUri } : {}),
        ...(anchor.xmlPath ? { xmlPath: anchor.xmlPath } : {}),
        role: "shape",
      },
    ],
  };
}
function findContainer(root: ContainerNode, id: string): ContainerNode | undefined {
  if (root.id === id) return root;
  for (const child of root.children)
    if (child.type === "container") {
      const found = findContainer(child, id);
      if (found) return found;
    }
  return undefined;
}

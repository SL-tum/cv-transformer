import type { ExtractionDecision, ExtractionProposal, TemplateCandidate } from "./candidate.js";

export function detectPptxStructure(
  candidates: TemplateCandidate[],
  acceptConfidence = 0.9,
  reviewConfidence = 0.7,
): { proposals: ExtractionProposal[]; decisions: ExtractionDecision[] } {
  const titles = candidates.filter((item) => item.nativeRole === "shapeGroupTitle");
  const fields: ExtractionProposal[] = [];
  const groups = groupBy(
    candidates.filter((candidate) => candidate.nativeRole !== "shapeGroupTitle"),
    (candidate) =>
      `${candidate.layoutFeatures?.slideIndex}:${candidate.layoutFeatures?.containerNodeId}`,
  );
  for (const group of groups.values()) {
    group.sort((a, b) => (a.layoutFeatures?.order ?? 0) - (b.layoutFeatures?.order ?? 0));
    const pairWithPlaceholder = group.length === 2 && Boolean(group[1]?.signals.text?.placeholder);
    const alternating =
      (group.length >= 4 || pairWithPlaceholder) &&
      group.length % 2 === 0 &&
      group
        .filter((_candidate, index) => index % 2 === 0)
        .every((candidate) => candidate.signals.text?.short && !candidate.signals.text.placeholder);
    if (!alternating) continue;
    for (let index = 0; index < group.length; index += 2) {
      const label = group[index]!;
      const value = group[index + 1]!;
      const placeholder = Boolean(value.signals.text?.placeholder);
      const confidence = placeholder
        ? 0.97
        : group.length >= 4 && label.styleFeatures.emphasized
          ? 0.94
          : group.length >= 4
            ? 0.91
            : 0.82;
      fields.push({
        id: uniqueId(slug(label.text ?? "field"), fields),
        type: listLike(label.text) ? "list" : "text",
        label: cleanLabel(label.text ?? "Field"),
        candidateIds: [label.id, value.id],
        valueCandidateIds: [value.id],
        confidence,
        source: "templateLabel",
        evidence: [
          {
            signal: "text",
            description: `Short template label '${label.text}' followed by content`,
          },
          {
            signal: "neighborhood",
            description: "Label and value are consecutive paragraphs in the same text shape",
          },
          {
            signal: "geometry",
            description: "Label and value share one Office shape and alignment",
          },
          ...(placeholder
            ? [
                {
                  signal: "text" as const,
                  description: "Value uses a recognizable template placeholder",
                },
              ]
            : []),
        ],
      });
    }
  }
  const proposals: ExtractionProposal[] = [];
  const consumed = new Set(fields.flatMap((field) => field.candidateIds));
  for (const title of titles) {
    const sectionId = uniqueId(slug(title.text ?? "section"), [...proposals, ...fields]);
    const label = cleanLabel(title.text ?? "Section");
    const evidence = [
      {
        signal: "nativeRole" as const,
        description: "Short title is the only direct text inside an Office shape group",
      },
      {
        signal: "geometry" as const,
        description: "Group also contains non-title layout or decorative elements",
      },
      { signal: "text" as const, description: `Group title is '${title.text}'` },
    ];
    proposals.push({
      id: sectionId,
      type: "container",
      label,
      candidateIds: [title.id],
      valueCandidateIds: [title.id],
      confidence: 0.93,
      source: "structuralRule",
      evidence,
    });
    const region = sectionRegion(title, titles);
    const children = fields.filter((field) =>
      field.candidateIds.some((id) => {
        const candidate = candidates.find((item) => item.id === id);
        return candidate ? inRegion(candidate, region) : false;
      }),
    );
    for (const child of children) child.parentId = sectionId;
    const placeholder = candidates
      .filter(
        (candidate) =>
          !consumed.has(candidate.id) &&
          candidate.signals.text?.placeholder &&
          inRegion(candidate, region),
      )
      .sort((a, b) => (a.layoutFeatures?.y ?? 0) - (b.layoutFeatures?.y ?? 0))[0];
    if (placeholder)
      proposals.push({
        id: `${sectionId}-content`,
        type: summaryLike(label) ? "text" : "list",
        label,
        candidateIds: [title.id, placeholder.id],
        valueCandidateIds: [placeholder.id],
        confidence: 0.97,
        source: "structuralRule",
        evidence: [
          ...evidence,
          {
            signal: "geometry",
            description:
              "Existing placeholder lies below this title and before the next section title",
          },
          {
            signal: "text",
            description: "Existing placeholder is used instead of inserting a new shape",
          },
        ],
        parentId: sectionId,
      });
    else if (!children.length && !hasOccupiedContent(candidates, title, region))
      proposals.push({
        id: `${sectionId}-content`,
        type: summaryLike(label) ? "text" : "list",
        label,
        candidateIds: [title.id],
        valueCandidateIds: [title.id],
        confidence: 0.9,
        source: "structuralRule",
        evidence: [
          ...evidence,
          {
            signal: "geometry",
            description:
              "Section region has no placeholder, child field, or content shape; insert a new shape as fallback",
          },
        ],
        parentId: sectionId,
        emptySection: true,
        insertionLayout: emptySectionLayout(title, titles),
      });
  }
  proposals.push(...fields);
  return { proposals, decisions: decide(proposals, acceptConfidence, reviewConfidence) };
}

export function detectDocxStructure(
  candidates: TemplateCandidate[],
  acceptConfidence = 0.9,
  reviewConfidence = 0.7,
): { proposals: ExtractionProposal[]; decisions: ExtractionDecision[] } {
  const proposals: ExtractionProposal[] = [];
  const groups = groupBy(
    candidates,
    (candidate) => candidate.layoutFeatures?.containerNodeId ?? "root",
  );
  for (const group of groups.values()) {
    group.sort((a, b) => (a.layoutFeatures?.order ?? 0) - (b.layoutFeatures?.order ?? 0));
    for (let i = 0; i < group.length - 1; i++) {
      const label = group[i]!;
      const value = group[i + 1]!;
      if (
        !label.signals.text?.short ||
        label.signals.text.placeholder ||
        value.nativeRole === "heading"
      )
        continue;
      const emphasized = label.styleFeatures.emphasized || label.nativeRole === "heading";
      const confidence = emphasized ? 0.94 : value.signals.text?.placeholder ? 0.9 : 0.72;
      proposals.push({
        id: uniqueId(slug(label.text ?? "field"), proposals),
        type: listLike(label.text) || value.nativeRole === "listParagraph" ? "list" : "text",
        label: cleanLabel(label.text ?? "Field"),
        candidateIds: [label.id, value.id],
        valueCandidateIds: [value.id],
        confidence,
        source: emphasized ? "structuralRule" : "templateLabel",
        evidence: [
          { signal: "text", description: `Short paragraph '${label.text}' precedes content` },
          {
            signal: "neighborhood",
            description: "Label and value are consecutive Word paragraphs",
          },
          ...(emphasized
            ? [
                {
                  signal: "style" as const,
                  description: "Label paragraph is visually emphasized or uses a heading style",
                },
              ]
            : []),
        ],
      });
      i++;
    }
  }
  return { proposals, decisions: decide(proposals, acceptConfidence, reviewConfidence) };
}

function decide(
  proposals: ExtractionProposal[],
  accept: number,
  review: number,
): ExtractionDecision[] {
  return proposals.map((proposal) => ({
    proposal,
    status:
      proposal.confidence >= accept
        ? "accepted"
        : proposal.confidence >= review
          ? "needsReview"
          : "rejected",
  }));
}
function cleanLabel(value: string) {
  return value.trim().replace(/:\s*$/, "");
}
function slug(value: string) {
  return (
    cleanLabel(value)
      .toLowerCase()
      .normalize("NFKD")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "field"
  );
}
function uniqueId(base: string, proposals: ExtractionProposal[]) {
  let id = base;
  let sequence = 2;
  while (proposals.some((proposal) => proposal.id === id)) id = `${base}-${sequence++}`;
  return id;
}
function listLike(label?: string) {
  return /(?:skills?|certificates?|languages?|responsibilities|achievements|qualifications|experience)/i.test(
    label ?? "",
  );
}
function summaryLike(label?: string) {
  return /(?:summary|profile|overview|introduction)/i.test(label ?? "");
}
function emptySectionLayout(candidate: TemplateCandidate, titles: TemplateCandidate[]) {
  const layout = candidate.layoutFeatures;
  const x = layout?.x ?? 0;
  const y = layout?.y ?? 0;
  const width = layout?.width ?? 3000000;
  const titleHeight = layout?.height ?? 300000;
  const top = y + titleHeight + 80000;
  const right = x + width;
  const next = titles
    .filter(
      (other) =>
        other !== candidate &&
        other.layoutFeatures?.slideIndex === layout?.slideIndex &&
        (other.layoutFeatures?.y ?? 0) > y &&
        overlap(
          x,
          right,
          other.layoutFeatures?.x ?? 0,
          (other.layoutFeatures?.x ?? 0) + (other.layoutFeatures?.width ?? 0),
        ) >=
          Math.min(width, other.layoutFeatures?.width ?? width) * 0.5,
    )
    .sort((a, b) => (a.layoutFeatures?.y ?? 0) - (b.layoutFeatures?.y ?? 0))[0];
  const bottom = next ? (next.layoutFeatures?.y ?? top + 1000000) - 80000 : top + 900000;
  return {
    x: x + 425000,
    y: top,
    width: Math.max(500000, width - 500000),
    height: Math.max(300000, bottom - top),
  };
}
function sectionRegion(candidate: TemplateCandidate, titles: TemplateCandidate[]) {
  const layout = candidate.layoutFeatures;
  const x = layout?.x ?? 0;
  const y = layout?.y ?? 0;
  const width = layout?.width ?? 3000000;
  const next = titles
    .filter(
      (other) =>
        other !== candidate &&
        other.layoutFeatures?.slideIndex === layout?.slideIndex &&
        (other.layoutFeatures?.y ?? 0) > y &&
        overlap(
          x,
          x + width,
          other.layoutFeatures?.x ?? 0,
          (other.layoutFeatures?.x ?? 0) + (other.layoutFeatures?.width ?? 0),
        ) >=
          Math.min(width, other.layoutFeatures?.width ?? width) * 0.5,
    )
    .sort((a, b) => (a.layoutFeatures?.y ?? 0) - (b.layoutFeatures?.y ?? 0))[0];
  return {
    slideIndex: layout?.slideIndex,
    x,
    right: x + width,
    top: y + (layout?.height ?? 0) * 0.5,
    bottom: next?.layoutFeatures?.y ?? Number.POSITIVE_INFINITY,
  };
}
function inRegion(
  candidate: TemplateCandidate,
  region: { slideIndex: number | undefined; x: number; right: number; top: number; bottom: number },
) {
  const layout = candidate.layoutFeatures;
  if (layout?.slideIndex !== region.slideIndex) return false;
  const centerX = (layout?.x ?? 0) + (layout?.width ?? 0) / 2;
  const centerY = (layout?.y ?? 0) + (layout?.height ?? 0) / 2;
  return (
    centerX >= region.x &&
    centerX <= region.right &&
    centerY >= region.top &&
    centerY < region.bottom
  );
}
function hasOccupiedContent(
  candidates: TemplateCandidate[],
  title: TemplateCandidate,
  region: { slideIndex: number | undefined; x: number; right: number; top: number; bottom: number },
) {
  const normalized = cleanLabel(title.text ?? "").toLowerCase();
  return candidates.some(
    (candidate) =>
      candidate !== title &&
      candidate.nativeRole !== "shapeGroupTitle" &&
      cleanLabel(candidate.text ?? "").toLowerCase() !== normalized &&
      !candidate.signals.text?.placeholder &&
      inRegion(candidate, region),
  );
}
function overlap(a1: number, a2: number, b1: number, b2: number) {
  return Math.max(0, Math.min(a2, b2) - Math.max(a1, b1));
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

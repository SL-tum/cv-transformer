import type { ExtractionSource, MarkerKind } from "../markers/marker.js";

export interface StyleFeatures {
  bold?: boolean;
  fontSizePt?: number;
  color?: string;
  styleId?: string;
  emphasized?: boolean;
}
export interface LayoutFeatures {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  slideIndex?: number;
  containerNodeId?: string;
  order?: number;
}
export interface ExtractionSignals {
  explicitMarker?: { kind: MarkerKind; id: string };
  text?: { value: string; short: boolean; placeholder: boolean };
  style?: StyleFeatures;
  geometry?: LayoutFeatures;
  neighborhood?: { previousCandidateId?: string; nextCandidateId?: string };
  repetition?: { signature: string; occurrences: number };
  nativeRole?: { role: string };
}
export interface TemplateCandidate {
  id: string;
  sourceNodeIds: string[];
  text?: string;
  styleFeatures: StyleFeatures;
  layoutFeatures?: LayoutFeatures;
  nativeRole?: string;
  marker?: { kind: MarkerKind; id: string };
  partUri?: string;
  xmlPath?: string;
  signals: ExtractionSignals;
}
export interface ExtractionEvidence {
  signal:
    "marker" | "text" | "style" | "geometry" | "neighborhood" | "repetition" | "nativeRole" | "llm";
  description: string;
  weight?: number;
}
export interface ExtractionProposal {
  id: string;
  type: "container" | "text" | "list" | "fieldGroup" | "collection" | "image";
  label?: string;
  candidateIds: string[];
  valueCandidateIds?: string[];
  confidence: number;
  source: ExtractionSource;
  evidence: ExtractionEvidence[];
  parentId?: string;
  emptySection?: boolean;
  insertionLayout?: { x: number; y: number; width: number; height: number };
}
export interface ExtractionDecision {
  proposal: ExtractionProposal;
  status: "accepted" | "needsReview" | "rejected";
}
export interface CandidateExtractionResult {
  candidates: TemplateCandidate[];
  proposals: ExtractionProposal[];
  decisions: ExtractionDecision[];
}

export type DocumentFormat = "docx" | "pptx";

export interface Length {
  value: number;
  unit: "emu" | "pt" | "px" | "twip" | "percent";
}

export interface Measurement {
  resolvedEmu: bigint;
  originalValue?: number;
  originalUnit?: string;
}

export interface Size {
  width: Length;
  height: Length;
}

export type ThemeColor =
  | "dark1"
  | "dark2"
  | "light1"
  | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink";

export interface ColorTransform {
  type: "tint" | "shade" | "alpha" | "lumMod" | "lumOff" | "satMod";
  value: number;
}

export type Color =
  | { type: "rgb"; value: string; alpha?: number }
  | { type: "scheme"; value: ThemeColor; transforms?: ColorTransform[] }
  | { type: "system"; value: string; fallback?: string }
  | { type: "preset"; value: string }
  | { type: "automatic" };

export interface ResolvedColor {
  original: Color;
  computedRgb: string;
}

export interface Transform {
  x?: Length;
  y?: Length;
  width?: Length;
  height?: Length;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  anchor?: Record<string, unknown>;
}

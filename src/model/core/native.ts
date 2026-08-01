export interface NativePayload {
  namespace?: string;
  qualifiedName?: string;
  attributes?: Record<string, string>;
  rawXml?: string;
  unknownChildren?: NativePayload[];
}

export interface NativeRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "Internal" | "External";
}

export interface NativePart {
  uri: string;
  contentType: string;
  relationshipType?: string;
  xml?: string;
  /** Exact uncompressed bytes. Kept even for XML so opaque parts round-trip. */
  data?: Uint8Array;
  binaryResourceId?: string;
  relationships: NativeRelationship[];
  parsedState: "fullyParsed" | "partiallyParsed" | "opaque";
}

export interface NativeStore {
  parts: Record<string, NativePart>;
  contentTypes?: {
    defaults: Record<string, string>;
    overrides: Record<string, string>;
    rawXml?: string;
  };
  rootRelationships?: NativeRelationship[];
  packageProperties?: Record<string, unknown>;
}

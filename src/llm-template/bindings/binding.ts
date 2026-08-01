export type BindingRepresentation = "plainText" | "joinedRuns" | "paragraphList" | "commaSeparatedText" | "shapeText" | "tableFields";
export interface ReadStrategy { type: "collectText" | "collectParagraphs" | "splitText"; separator?: string }
export interface WriteStrategy { type: "replaceTextPreservingRuns" | "replaceParagraphList" | "replaceShapeText" | "clonePrototype" }
export interface FormatStrategy { type: "preserveExisting" | "copyPrototype" }
export interface SourceLocation { nodeId: string; partUri?: string; xmlPath?: string; role?: "root" | "text" | "paragraph" | "shape" | "field" }
export interface TemplateBinding { templateNodeId: string; sourceNodeIds: string[]; sourceFormat: "docx" | "pptx"; representation: BindingRepresentation; readStrategy: ReadStrategy; writeStrategy: WriteStrategy; formatStrategy: FormatStrategy; /** Internal only; omitted from LLM JSON. */ locations?: SourceLocation[] }
export interface PrototypeFieldBinding { fieldId: string; markerId: string; relativeSourceNodeIds: string[] }
export interface PrototypeBinding { id: string; collectionId: string; rootNodeIds: string[]; cloneStrategy: "deepCloneParagraphRange" | "deepCloneTableRow" | "deepCloneShapeGroup"; insertionAnchorNodeId: string; insertionPosition: "before" | "after" | "inside"; regenerateNodeIds: boolean; regenerateOfficeIds: boolean; preserveStyles: boolean; clearEditableContent: boolean; fieldBindings: Record<string, PrototypeFieldBinding>; locations?: SourceLocation[]; nativeXml?: string }

import type { ContainerNode } from "./template-node.js";
export interface LlmTemplateDocument { documentId: string; revision: number; sourceFormat: "docx" | "pptx"; root: ContainerNode }

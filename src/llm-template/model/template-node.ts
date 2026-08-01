import type { ContentConstraints } from "./constraints.js";

export interface TemplateNodeBase { id: string; type: string; label?: string; editable: boolean; required?: boolean; constraints?: ContentConstraints; metadata?: Record<string, unknown> }
export interface ContainerNode extends TemplateNodeBase { type: "container"; children: TemplateNode[] }
export interface TextNode extends TemplateNodeBase { type: "text"; value: string; placeholder?: string }
export interface ListItem { id: string; value: string }
export interface ListNode extends TemplateNodeBase { type: "list"; items: ListItem[]; repeatable: boolean }
export interface TemplateField { id: string; key: string; label?: string; type: "text" | "list" | "image"; value?: string; items?: ListItem[]; editable: boolean; constraints?: ContentConstraints }
export interface FieldGroupNode extends TemplateNodeBase { type: "fieldGroup"; fields: TemplateField[] }
export interface CollectionNode extends TemplateNodeBase { type: "collection"; repeatable: true; items: FieldGroupNode[]; prototypeId: string; minItems?: number; maxItems?: number }
export interface ImageNode extends TemplateNodeBase { type: "image"; description?: string; currentResourceId?: string; allowedMimeTypes?: string[]; replaceable: boolean }
export type TemplateNode = ContainerNode | TextNode | ListNode | FieldGroupNode | CollectionNode | ImageNode;

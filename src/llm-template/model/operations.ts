export interface SetTextOperation { op: "setText"; targetId: string; value: string }
export interface SetListOperation { op: "setList"; targetId: string; items: string[] }
export interface AddListItemOperation { op: "appendListItem"; targetId: string; value: string }
export interface RemoveListItemOperation { op: "removeListItem"; targetId: string; itemId?: string; index?: number }
export interface AddCollectionItemOperation { op: "appendCollectionItem"; targetId: string; value: Record<string, string | string[]> }
export interface UpdateCollectionItemOperation { op: "updateCollectionItem"; targetId: string; itemId: string; value: Record<string, string | string[]> }
export interface RemoveCollectionItemOperation { op: "removeCollectionItem"; targetId: string; itemId: string }
export type TemplateOperation = SetTextOperation | SetListOperation | AddListItemOperation | RemoveListItemOperation | AddCollectionItemOperation | UpdateCollectionItemOperation | RemoveCollectionItemOperation;
export interface TemplateOperationBatch { documentId: string; revision: number; operations: TemplateOperation[] }

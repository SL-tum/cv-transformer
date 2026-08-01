export interface ContentConstraints {
  hard?: { maxCharacters?: number; maxItems?: number; allowedOperations?: string[] };
  soft?: { recommendedCharacters?: number; recommendedWords?: number; recommendedWordsPerItem?: number; recommendedItems?: number; maxLines?: number; overflowRisk?: "low" | "medium" | "high" };
}
export interface ConstraintWarning { code: string; targetId: string; message: string }

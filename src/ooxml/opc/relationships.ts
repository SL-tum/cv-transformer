import type { NativeRelationship } from "../../model/core/native.js";
import { attr, elementsByLocalName, parseXml } from "../xml.js";
import { normalizePartUri } from "./content-types.js";

export function parseRelationships(xml: string): NativeRelationship[] {
  return elementsByLocalName(parseXml(xml), "Relationship").map((element) => {
    const id = attr(element, "Id");
    const type = attr(element, "Type");
    const target = attr(element, "Target");
    if (!id || !type || !target) throw new Error("Relationship requires Id, Type, and Target");
    const mode = attr(element, "TargetMode");
    return {
      id,
      type,
      target,
      ...(mode === "External" ? { targetMode: "External" as const } : {}),
    };
  });
}

export function relationshipPartUri(sourcePartUri?: string): string {
  if (!sourcePartUri || sourcePartUri === "/") return "/_rels/.rels";
  const normalized = normalizePartUri(sourcePartUri);
  const slash = normalized.lastIndexOf("/");
  return `${normalized.slice(0, slash)}/_rels/${normalized.slice(slash + 1)}.rels`;
}

export function resolveRelationshipTarget(
  sourcePartUri: string | undefined,
  relationship: NativeRelationship,
): string | undefined {
  if (relationship.targetMode === "External") return undefined;
  const base =
    sourcePartUri && sourcePartUri !== "/"
      ? sourcePartUri.slice(0, sourcePartUri.lastIndexOf("/") + 1)
      : "/";
  const segments = `${base}${relationship.target}`.split("/");
  const resolved: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") resolved.pop();
    else resolved.push(segment);
  }
  return `/${resolved.join("/")}`;
}

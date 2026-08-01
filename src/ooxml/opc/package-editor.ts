import { strToU8 } from "fflate";
import type { NativePart, NativeRelationship } from "../../model/core/native.js";
import { parseXml } from "../xml.js";
import { ensureContentType, normalizePartUri, removeContentType } from "./content-types.js";
import type { OpcPackage } from "./package.js";
import { relationshipPartUri } from "./relationships.js";

const RELS_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
export interface AddPartOptions {
  preferredUri: string;
  contentType: string;
  data: Uint8Array | string;
  parsedState?: NativePart["parsedState"];
  preferDefaultContentType?: boolean;
}

export function addPart(pkg: OpcPackage, options: AddPartOptions): NativePart {
  const uri = nextAvailablePartUri(pkg, options.preferredUri);
  const data = typeof options.data === "string" ? strToU8(options.data) : options.data;
  const part: NativePart = {
    uri,
    contentType: options.contentType,
    data,
    ...(typeof options.data === "string" ? { xml: options.data } : {}),
    relationships: [],
    parsedState: options.parsedState ?? "opaque",
  };
  pkg.parts.set(uri, part);
  ensureContentType(pkg.contentTypes, uri, options.contentType, options.preferDefaultContentType);
  return part;
}

export function removePart(pkg: OpcPackage, partUri: string): void {
  const uri = normalizePartUri(partUri);
  pkg.parts.delete(uri);
  removeContentType(pkg.contentTypes, uri);
  pkg.parts.delete(relationshipPartUri(uri));
}

export function addRelationship(
  pkg: OpcPackage,
  sourcePartUri: string | undefined,
  type: string,
  target: string,
  targetMode?: "Internal" | "External",
): NativeRelationship {
  const relationships = sourcePartUri
    ? requirePart(pkg, sourcePartUri).relationships
    : pkg.rootRelationships;
  const used = new Set(relationships.map((item) => item.id));
  let sequence = 1;
  while (used.has(`rId${sequence}`)) sequence++;
  const relationship: NativeRelationship = {
    id: `rId${sequence}`,
    type,
    target,
    ...(targetMode ? { targetMode } : {}),
  };
  relationships.push(relationship);
  syncRelationshipsPart(pkg, sourcePartUri);
  return relationship;
}

export function removeRelationship(
  pkg: OpcPackage,
  sourcePartUri: string | undefined,
  relationshipId: string,
): boolean {
  const relationships = sourcePartUri
    ? requirePart(pkg, sourcePartUri).relationships
    : pkg.rootRelationships;
  const index = relationships.findIndex((item) => item.id === relationshipId);
  if (index < 0) return false;
  relationships.splice(index, 1);
  syncRelationshipsPart(pkg, sourcePartUri);
  return true;
}

export function garbageCollectRelationships(
  pkg: OpcPackage,
  sourcePartUri: string,
): NativeRelationship[] {
  const part = requirePart(pkg, sourcePartUri);
  if (!part.xml) return [];
  const used = referencedRelationshipIds(part.xml);
  const removed = part.relationships.filter((relationship) => !used.has(relationship.id));
  part.relationships = part.relationships.filter((relationship) => used.has(relationship.id));
  if (removed.length) syncRelationshipsPart(pkg, sourcePartUri);
  return removed;
}

export interface AddResourceOptions {
  sourcePartUri: string;
  directory: string;
  fileName: string;
  contentType: string;
  relationshipType: string;
  data: Uint8Array;
  preferDefaultContentType?: boolean;
}
export function addResourcePart(
  pkg: OpcPackage,
  options: AddResourceOptions,
): { part: NativePart; relationship: NativeRelationship } {
  const directory = normalizePartUri(options.directory).replace(/\/$/, "");
  const part = addPart(pkg, {
    preferredUri: `${directory}/${options.fileName}`,
    contentType: options.contentType,
    data: options.data,
    preferDefaultContentType: options.preferDefaultContentType ?? true,
  });
  const sourceDirectory = options.sourcePartUri.slice(
    0,
    options.sourcePartUri.lastIndexOf("/") + 1,
  );
  const target = relativeTarget(sourceDirectory, part.uri);
  return {
    part,
    relationship: addRelationship(pkg, options.sourcePartUri, options.relationshipType, target),
  };
}

export function nextAvailablePartUri(pkg: OpcPackage, preferredUri: string): string {
  const normalized = normalizePartUri(preferredUri);
  if (!pkg.parts.has(normalized)) return normalized;
  const dot = normalized.lastIndexOf(".");
  const stem = dot > normalized.lastIndexOf("/") ? normalized.slice(0, dot) : normalized;
  const extension = dot > normalized.lastIndexOf("/") ? normalized.slice(dot) : "";
  let index = 2;
  while (pkg.parts.has(`${stem}${index}${extension}`)) index++;
  return `${stem}${index}${extension}`;
}

function syncRelationshipsPart(pkg: OpcPackage, sourcePartUri?: string): void {
  const relationships = sourcePartUri
    ? requirePart(pkg, sourcePartUri).relationships
    : pkg.rootRelationships;
  const uri = relationshipPartUri(sourcePartUri);
  const xml = relationshipsXml(relationships);
  const existing = pkg.parts.get(uri);
  if (existing) {
    existing.xml = xml;
    existing.data = strToU8(xml);
  } else {
    pkg.parts.set(uri, {
      uri,
      contentType: RELS_TYPE,
      xml,
      data: strToU8(xml),
      relationships: [],
      parsedState: "fullyParsed",
    });
  }
  ensureContentType(pkg.contentTypes, uri, RELS_TYPE, true);
}
function relationshipsXml(relationships: NativeRelationship[]): string {
  const escape = (x: string) =>
    x.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((item) => `<Relationship Id="${escape(item.id)}" Type="${escape(item.type)}" Target="${escape(item.target)}"${item.targetMode ? ` TargetMode="${item.targetMode}"` : ""}/>`).join("")}</Relationships>`;
}
function referencedRelationshipIds(xml: string): Set<string> {
  const document = parseXml(xml);
  const result = new Set<string>();
  const all = document.getElementsByTagName("*");
  for (let i = 0; i < all.length; i++) {
    const element = all.item(i);
    if (!element) continue;
    for (let j = 0; j < element.attributes.length; j++) {
      const attribute = element.attributes.item(j);
      if (
        attribute &&
        (attribute.prefix === "r" || attribute.namespaceURI?.includes("relationships")) &&
        ["id", "embed", "link"].includes(
          attribute.localName ?? attribute.name.split(":").at(-1) ?? "",
        )
      )
        result.add(attribute.value);
    }
  }
  return result;
}
function requirePart(pkg: OpcPackage, uri: string): NativePart {
  const part = pkg.parts.get(normalizePartUri(uri));
  if (!part) throw new Error(`Part not found: ${uri}`);
  return part;
}
function relativeTarget(sourceDirectory: string, targetUri: string): string {
  const from = sourceDirectory.split("/").filter(Boolean);
  const to = targetUri.split("/").filter(Boolean);
  while (from.length && to.length && from[0] === to[0]) {
    from.shift();
    to.shift();
  }
  return `${"../".repeat(from.length)}${to.join("/")}`;
}

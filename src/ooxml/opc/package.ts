import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { NativePart, NativeRelationship, NativeStore } from "../../model/core/native.js";
import { contentTypeFor, normalizePartUri, parseContentTypes } from "./content-types.js";
import { parseRelationships, relationshipPartUri } from "./relationships.js";

export interface OpcPackage { contentTypes: ReturnType<typeof parseContentTypes>; parts: Map<string, NativePart>; rootRelationships: NativeRelationship[] }

const isXmlType = (type: string, uri: string) => type.includes("xml") || uri.endsWith(".xml") || uri.endsWith(".rels");

export function loadOpcPackage(input: Uint8Array): OpcPackage {
  const entries = unzipSync(input);
  const contentTypesBytes = entries["[Content_Types].xml"];
  if (!contentTypesBytes) throw new Error("Not an OPC package: [Content_Types].xml is missing");
  const contentTypes = parseContentTypes(strFromU8(contentTypesBytes));
  const parts = new Map<string, NativePart>();
  for (const [name, data] of Object.entries(entries)) {
    if (name.endsWith("/") || name === "[Content_Types].xml") continue;
    const uri = normalizePartUri(name); const contentType = contentTypeFor(uri, contentTypes);
    parts.set(uri, { uri, contentType, data, ...(isXmlType(contentType, uri) ? { xml: strFromU8(data) } : {}), relationships: [], parsedState: "opaque" });
  }
  const rootRelationshipPart = parts.get("/_rels/.rels");
  const rootRelationships = rootRelationshipPart?.xml ? parseRelationships(rootRelationshipPart.xml) : [];
  for (const part of parts.values()) {
    if (part.uri.endsWith(".rels") || part.uri === "/_rels/.rels") continue;
    const rels = parts.get(relationshipPartUri(part.uri));
    if (rels?.xml) part.relationships = parseRelationships(rels.xml);
  }
  return { contentTypes, parts, rootRelationships };
}

export function writeOpcPackage(pkg: OpcPackage, options: { level?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 } = {}): Uint8Array {
  const entries: Record<string, Uint8Array> = {};
  entries["[Content_Types].xml"] = strToU8(pkg.contentTypes.rawXml ?? buildContentTypesXml(pkg.contentTypes));
  for (const part of pkg.parts.values()) {
    const unchangedXml = part.xml !== undefined && part.data !== undefined && strFromU8(part.data) === part.xml;
    entries[part.uri.replace(/^\//, "")] = unchangedXml ? part.data! : part.xml !== undefined ? strToU8(part.xml) : part.data ?? new Uint8Array();
  }
  return zipSync(entries, { level: options.level ?? 6 });
}

export function packageToNativeStore(pkg: OpcPackage): NativeStore {
  return { parts: Object.fromEntries(pkg.parts), contentTypes: pkg.contentTypes, rootRelationships: pkg.rootRelationships };
}
export function nativeStoreToPackage(store: NativeStore): OpcPackage {
  if (!store.contentTypes) throw new Error("NativeStore has no content types");
  return { parts: new Map(Object.entries(store.parts)), contentTypes: store.contentTypes, rootRelationships: store.rootRelationships ?? [] };
}
export const cloneOpcPackage = (pkg: OpcPackage): OpcPackage => loadOpcPackage(writeOpcPackage(pkg, { level: 0 }));

export function validateOpcPackage(pkg: OpcPackage): string[] {
  const issues: string[] = [];
  if (!pkg.contentTypes.rawXml && !Object.keys(pkg.contentTypes.defaults).length && !Object.keys(pkg.contentTypes.overrides).length) issues.push("Content types are empty");
  const validateRelationships = (source: string | undefined, relationships: NativeRelationship[]) => { for (const relationship of relationships) { if (relationship.targetMode === "External") continue; const target = resolveTarget(source, relationship.target); if (!pkg.parts.has(target)) issues.push(`Missing relationship target ${target} from ${source ?? "/"} (${relationship.id})`); } };
  validateRelationships(undefined, pkg.rootRelationships);
  for (const part of pkg.parts.values()) validateRelationships(part.uri, part.relationships);
  return issues;
}

function resolveTarget(source: string | undefined, target: string): string { const base = source ? source.slice(0, source.lastIndexOf("/") + 1) : "/"; const output: string[] = []; for (const segment of `${base}${target}`.split("/")) { if (!segment || segment === ".") continue; if (segment === "..") output.pop(); else output.push(segment); } return `/${output.join("/")}`; }

function buildContentTypesXml(types: ReturnType<typeof parseContentTypes>): string {
  const escape = (value: string) => value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
  const defaults = Object.entries(types.defaults).map(([extension, type]) => `<Default Extension="${escape(extension)}" ContentType="${escape(type)}"/>`).join("");
  const overrides = Object.entries(types.overrides).map(([name, type]) => `<Override PartName="${escape(name)}" ContentType="${escape(type)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
}

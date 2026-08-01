import { attr, elementsByLocalName, parseXml } from "../xml.js";

export interface ContentTypes { defaults: Record<string, string>; overrides: Record<string, string>; rawXml?: string }

export function parseContentTypes(xml: string): ContentTypes {
  const document = parseXml(xml);
  const defaults: Record<string, string> = {};
  const overrides: Record<string, string> = {};
  for (const item of elementsByLocalName(document, "Default")) {
    const extension = attr(item, "Extension"); const type = attr(item, "ContentType");
    if (extension && type) defaults[extension.toLowerCase()] = type;
  }
  for (const item of elementsByLocalName(document, "Override")) {
    const name = attr(item, "PartName"); const type = attr(item, "ContentType");
    if (name && type) overrides[normalizePartUri(name)] = type;
  }
  return { defaults, overrides, rawXml: xml };
}

export function contentTypeFor(uri: string, types: ContentTypes): string {
  const normalized = normalizePartUri(uri);
  const override = types.overrides[normalized];
  if (override) return override;
  const extension = normalized.split(".").at(-1)?.toLowerCase();
  return extension ? types.defaults[extension] ?? "application/octet-stream" : "application/octet-stream";
}
export const normalizePartUri = (uri: string): string => `/${uri.replace(/^\/+/, "")}`;

export function ensureContentType(types: ContentTypes, partUri: string, contentType: string, preferDefault = false): void {
  const uri = normalizePartUri(partUri); const extension = uri.split(".").at(-1)?.toLowerCase();
  if (preferDefault && extension && (!types.defaults[extension] || types.defaults[extension] === contentType)) types.defaults[extension] = contentType;
  else if (!extension || types.defaults[extension] !== contentType) types.overrides[uri] = contentType;
  delete types.rawXml;
}
export function removeContentType(types: ContentTypes, partUri: string): void { delete types.overrides[normalizePartUri(partUri)]; delete types.rawXml; }

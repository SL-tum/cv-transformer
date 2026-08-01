import type { RichDocument } from "../../model/core/document.js";
import type { FlowDocumentRoot } from "../../model/word/index.js";
import { attr, elementPath, elementsByLocalName, parseXml } from "../../ooxml/xml.js";
import type { TemplateMarker } from "./marker.js";
import { parseMarker } from "./marker.js";
import { indexRdtNodes, sourceKey } from "./common.js";

export function readDocxMarkers(document: RichDocument<FlowDocumentRoot>): TemplateMarker[] {
  if (!document.nativeStore) throw new Error("Marker extraction requires NativeStore"); const index = indexRdtNodes(document); const markers: TemplateMarker[] = [];
  for (const part of Object.values(document.nativeStore.parts)) {
    if (!part.xml || !/wordprocessingml|header\+xml|footer\+xml/.test(part.contentType)) continue; const dom = parseXml(part.xml);
    for (const sdt of elementsByLocalName(dom, "sdt")) { const properties = elementsByLocalName(sdt, "sdtPr")[0]; const tag = properties && elementsByLocalName(properties, "tag")[0]; const alias = properties && elementsByLocalName(properties, "alias")[0]; const raw = tag && attr(tag, "val") || alias && attr(alias, "val"); if (!raw) continue; const parsed = parseMarker(raw); if (!parsed) continue; const content = elementsByLocalName(sdt, "sdtContent")[0] ?? sdt; markers.push(makeMarker(parsed.kind, parsed.id, raw, part.uri, sdt, content, index.bySource)); }
    const all = [...elementsByLocalName(dom, "bookmarkStart")]; for (const start of all) { const raw = attr(start, "name"); const bookmarkId = attr(start, "id"); if (!raw || !bookmarkId) continue; const parsed = parseMarker(raw); if (!parsed) continue; const nodes = documentOrder(dom.documentElement); const startIndex = nodes.indexOf(start); const endIndex = nodes.findIndex((element, i) => i > startIndex && element.localName === "bookmarkEnd" && attr(element, "id") === bookmarkId); const texts = nodes.slice(startIndex + 1, endIndex < 0 ? startIndex + 1 : endIndex).filter((element) => element.localName === "t"); markers.push(markerFromTexts(parsed.kind, parsed.id, raw, part.uri, start, texts, index.bySource)); }
  }
  return markers;
}
function makeMarker(kind: TemplateMarker["kind"], id: string, raw: string, partUri: string, root: Element, content: Element, index: Map<string, import("../../model/core/node.js").BaseNode>): TemplateMarker { return markerFromTexts(kind, id, raw, partUri, root, elementsByLocalName(content, "t"), index); }
function markerFromTexts(kind: TemplateMarker["kind"], id: string, raw: string, partUri: string, root: Element, texts: Element[], index: Map<string, import("../../model/core/node.js").BaseNode>): TemplateMarker { const textPaths = texts.map(elementPath); const textNodeIds = textPaths.flatMap((path) => { const node = index.get(sourceKey(partUri, path)); return node ? [node.id] : []; }); let ancestor: Element | null = root; let rootNodeId: string | undefined; while (ancestor && !rootNodeId) { rootNodeId = index.get(sourceKey(partUri, elementPath(ancestor)))?.id; ancestor = ancestor.parentNode?.nodeType === 1 ? ancestor.parentNode as Element : null; } return { raw, kind, id, source: "explicitMarker", partUri, xmlPath: elementPath(root), ...(rootNodeId ? { rootNodeId } : {}), textNodeIds, textPaths, nativeXml: root.toString() }; }
function documentOrder(root: Element): Element[] { const result: Element[] = []; const walk = (element: Element) => { result.push(element); for (let child = element.firstChild; child; child = child.nextSibling) if (child.nodeType === 1) walk(child as Element); }; walk(root); return result; }

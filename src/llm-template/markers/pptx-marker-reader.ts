import type { RichDocument } from "../../model/core/document.js";
import type { PresentationDocumentRoot } from "../../model/presentation/index.js";
import { attr, elementPath, elementsByLocalName, parseXml } from "../../ooxml/xml.js";
import type { TemplateMarker } from "./marker.js";
import { parseMarker } from "./marker.js";
import { indexRdtNodes, sourceKey } from "./common.js";

export function readPptxMarkers(
  document: RichDocument<PresentationDocumentRoot>,
): TemplateMarker[] {
  if (!document.nativeStore) throw new Error("Marker extraction requires NativeStore");
  const index = indexRdtNodes(document);
  const result: TemplateMarker[] = [];
  for (const part of Object.values(document.nativeStore.parts)) {
    if (!part.xml || !/presentationml\.(slide|slideLayout|slideMaster)\+xml/.test(part.contentType))
      continue;
    const dom = parseXml(part.xml);
    for (const properties of elementsByLocalName(dom, "cNvPr")) {
      const candidates = [
        attr(properties, "name"),
        attr(properties, "descr"),
        attr(properties, "title"),
      ].filter((x): x is string => Boolean(x));
      const raw = candidates.find((value) => parseMarker(value));
      if (!raw) continue;
      const parsed = parseMarker(raw)!;
      let shape = properties.parentNode?.parentNode as Element | undefined;
      while (shape && !["sp", "pic", "grpSp", "graphicFrame"].includes(shape.localName))
        shape = shape.parentNode?.nodeType === 1 ? (shape.parentNode as Element) : undefined;
      if (!shape) continue;
      const rootPath = elementPath(shape);
      const rootNode = index.bySource.get(sourceKey(part.uri, rootPath));
      const textPaths = elementsByLocalName(shape, "t").map(elementPath);
      const textNodeIds = textPaths.flatMap((path) => {
        const node = index.bySource.get(sourceKey(part.uri, path));
        return node ? [node.id] : [];
      });
      result.push({
        raw,
        kind: parsed.kind,
        id: parsed.id,
        source: "explicitMarker",
        partUri: part.uri,
        xmlPath: rootPath,
        ...(rootNode ? { rootNodeId: rootNode.id } : {}),
        textNodeIds,
        textPaths,
        nativeXml: shape.toString(),
      });
    }
  }
  return result;
}

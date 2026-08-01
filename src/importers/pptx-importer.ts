import { createRichDocument, type RichDocument } from "../model/core/document.js";
import type { Transform } from "../model/core/primitives.js";
import type { GraphicNode, GroupNode, PictureNode, ShapeNode } from "../model/drawing/index.js";
import type { ParagraphNode, RunProperties, TextBody, TextRunNode } from "../model/text/index.js";
import type {
  PlaceholderNode,
  PresentationDocumentRoot,
  SlideElementNode,
  SlideLayoutNode,
  SlideMasterNode,
  SlideNode,
} from "../model/presentation/index.js";
import { stableNodeId } from "../id.js";
import { loadOpcPackage, packageToNativeStore, type OpcPackage } from "../ooxml/opc/package.js";
import { resolveRelationshipTarget } from "../ooxml/opc/relationships.js";
import { attr, childElements, elementPath, elementsByLocalName, parseXml } from "../ooxml/xml.js";
import { mainPartUri, relatedPartUri } from "./common.js";

const makeId = (uri: string, element: Element, prefix: string) => {
  const nativeId = attr(elementsByLocalName(element, "cNvPr")[0] ?? element, "id");
  return stableNodeId(
    {
      sourceFormat: "pptx",
      partUri: uri,
      xmlPath: elementPath(element),
      ...(nativeId ? { nativeId } : {}),
    },
    prefix,
  );
};
const source = (uri: string, element: Element) => ({
  sourceFormat: "pptx" as const,
  partUri: uri,
  xmlPath: elementPath(element),
});

export function importPptx(input: Uint8Array): RichDocument<PresentationDocumentRoot> {
  const pkg = loadOpcPackage(input);
  const presentationUri = mainPartUri(pkg);
  const part = pkg.parts.get(presentationUri);
  if (!part?.xml) throw new Error(`PPTX presentation is missing: ${presentationUri}`);
  const dom = parseXml(part.xml);
  const slideSizeElement = elementsByLocalName(dom, "sldSz")[0];
  const slideUris = relationshipUrisFromIdList(pkg, presentationUri, dom, "sldId");
  const masterUris = relationshipUrisFromIdList(pkg, presentationUri, dom, "sldMasterId");
  const slideMasters = masterUris.map((uri) => parseMaster(pkg, uri));
  const layoutUris = [...new Set(slideMasters.flatMap((master) => master.layoutRefs))];
  const slideLayouts = layoutUris.map((uri) => parseLayout(pkg, uri));
  const slides = slideUris.map((uri) => parseSlide(pkg, uri));
  const root: PresentationDocumentRoot = {
    id: makeId(presentationUri, dom.documentElement, "presentation"),
    type: "presentation",
    slideSize: {
      width: {
        value: Number((slideSizeElement && attr(slideSizeElement, "cx")) || 0),
        unit: "emu",
      },
      height: {
        value: Number((slideSizeElement && attr(slideSizeElement, "cy")) || 0),
        unit: "emu",
      },
    },
    slideMasters,
    slideLayouts,
    slides,
    source: source(presentationUri, dom.documentElement),
  };
  const result = createRichDocument(`pptx:${root.id}`, root);
  result.nativeStore = packageToNativeStore(pkg);
  importThemes(result, pkg);
  return result;
}

function relationshipUrisFromIdList(
  pkg: OpcPackage,
  sourceUri: string,
  dom: Document,
  elementName: string,
): string[] {
  return elementsByLocalName(dom, elementName).flatMap((element) => {
    const id =
      element.getAttribute("r:id") ||
      element.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      );
    const uri = id && relatedPartUri(pkg, sourceUri, id);
    return uri ? [uri] : [];
  });
}
function parseMaster(pkg: OpcPackage, uri: string): SlideMasterNode {
  const dom = xmlPart(pkg, uri);
  const layoutRefs = relationshipUrisFromIdList(pkg, uri, dom, "sldLayoutId");
  const themeRelationship = pkg.parts
    .get(uri)
    ?.relationships.find((rel) => rel.type.endsWith("/theme"));
  const themeRef = themeRelationship && resolveRelationshipTarget(uri, themeRelationship);
  return {
    id: uri,
    type: "slideMaster",
    shapes: parseShapeTree(dom, uri),
    layoutRefs,
    ...(themeRef ? { themeRef } : {}),
    source: source(uri, dom.documentElement),
  };
}
function parseLayout(pkg: OpcPackage, uri: string): SlideLayoutNode {
  const dom = xmlPart(pkg, uri);
  const masterRelationship = pkg.parts
    .get(uri)
    ?.relationships.find((rel) => rel.type.endsWith("/slideMaster"));
  const masterRef = masterRelationship && resolveRelationshipTarget(uri, masterRelationship);
  return {
    id: uri,
    type: "slideLayout",
    shapes: parseShapeTree(dom, uri),
    ...(masterRef ? { masterRef } : {}),
    source: source(uri, dom.documentElement),
  };
}
function parseSlide(pkg: OpcPackage, uri: string): SlideNode {
  const dom = xmlPart(pkg, uri);
  const layoutRelationship = pkg.parts
    .get(uri)
    ?.relationships.find((rel) => rel.type.endsWith("/slideLayout"));
  const layoutRef = layoutRelationship && resolveRelationshipTarget(uri, layoutRelationship);
  return {
    id: uri,
    type: "slide",
    shapes: parseShapeTree(dom, uri),
    ...(layoutRef ? { layoutRef } : {}),
    source: source(uri, dom.documentElement),
  };
}
function xmlPart(pkg: OpcPackage, uri: string): Document {
  const xml = pkg.parts.get(uri)?.xml;
  if (!xml) throw new Error(`XML part is missing: ${uri}`);
  return parseXml(xml);
}

function parseShapeTree(dom: Document, uri: string): SlideElementNode[] {
  const tree = elementsByLocalName(dom, "spTree")[0];
  return tree ? parseShapeChildren(tree, uri) : [];
}
function parseShapeChildren(parent: Element, uri: string): SlideElementNode[] {
  const result: SlideElementNode[] = [];
  let zIndex = 0;
  for (const element of childElements(parent)) {
    let parsed: SlideElementNode | undefined;
    if (element.localName === "sp" || element.localName === "cxnSp")
      parsed = parseShape(element, uri);
    else if (element.localName === "pic") parsed = parsePicture(element, uri);
    else if (element.localName === "grpSp") parsed = parseGroup(element, uri);
    if (parsed) {
      parsed.metadata = { ...parsed.metadata, zIndex: zIndex++ };
      result.push(parsed);
    }
  }
  return result;
}
function parseShape(element: Element, uri: string): ShapeNode | PlaceholderNode {
  const placeholder = elementsByLocalName(element, "ph")[0];
  const preset = attr(elementsByLocalName(element, "prstGeom")[0] ?? element, "prst");
  const textBody = parseTextBody(element, uri);
  const base = {
    id: makeId(uri, element, placeholder ? "placeholder" : "shape"),
    geometry: { ...(preset ? { preset } : {}) },
    transform: parseTransform(element),
    ...(textBody ? { textBody } : {}),
    source: source(uri, element),
  };
  if (placeholder)
    return {
      ...base,
      type: "placeholder",
      placeholderType: normalizePlaceholder(attr(placeholder, "type")),
      ...(attr(placeholder, "idx") ? { placeholderIndex: Number(attr(placeholder, "idx")) } : {}),
    };
  return { ...base, type: "shape" };
}
function parsePicture(element: Element, uri: string): PictureNode {
  const blip = elementsByLocalName(element, "blip")[0];
  return {
    id: makeId(uri, element, "picture"),
    type: "picture",
    resourceId: attr(blip ?? element, "embed") ?? "",
    transform: parseTransform(element),
    source: source(uri, element),
  };
}
function parseGroup(element: Element, uri: string): GroupNode {
  return {
    id: makeId(uri, element, "group"),
    type: "group",
    transform: parseTransform(element),
    elements: parseShapeChildren(element, uri) as GraphicNode[],
    source: source(uri, element),
  };
}
function parseTransform(element: Element): Transform {
  const xfrm = elementsByLocalName(element, "xfrm")[0];
  const off = xfrm && childElements(xfrm, "off")[0];
  const ext = xfrm && childElements(xfrm, "ext")[0];
  return {
    ...(off
      ? {
          x: { value: Number(attr(off, "x") ?? 0), unit: "emu" },
          y: { value: Number(attr(off, "y") ?? 0), unit: "emu" },
        }
      : {}),
    ...(ext
      ? {
          width: { value: Number(attr(ext, "cx") ?? 0), unit: "emu" },
          height: { value: Number(attr(ext, "cy") ?? 0), unit: "emu" },
        }
      : {}),
    ...(xfrm && attr(xfrm, "rot") ? { rotation: Number(attr(xfrm, "rot")) / 60000 } : {}),
  };
}
function parseTextBody(element: Element, uri: string): TextBody | undefined {
  const txBody = childElements(element, "txBody")[0];
  if (!txBody) return undefined;
  return {
    paragraphs: childElements(txBody, "p").map((paragraph) => parseParagraph(paragraph, uri)),
  };
}
function parseParagraph(element: Element, uri: string): ParagraphNode {
  const runs: TextRunNode[] = [];
  for (const run of childElements(element)) {
    if (run.localName !== "r" && run.localName !== "fld") continue;
    const text = childElements(run, "t")[0];
    if (text)
      runs.push({
        id: makeId(uri, text, "text"),
        type: "textRun",
        text: text.textContent ?? "",
        properties: parseRunProperties(childElements(run, "rPr")[0]),
        source: source(uri, text),
      });
  }
  return {
    id: makeId(uri, element, "p"),
    type: "paragraph",
    runs,
    properties: {},
    source: source(uri, element),
  };
}
function parseRunProperties(element?: Element): RunProperties {
  if (!element) return {};
  const result: RunProperties = {};
  if (attr(element, "b") === "1") result.bold = true;
  if (attr(element, "i") === "1") result.italic = true;
  const size = attr(element, "sz");
  if (size) result.fontSize = { value: Number(size) / 100, unit: "pt" };
  return result;
}
function normalizePlaceholder(value?: string): PlaceholderNode["placeholderType"] {
  const map: Record<string, PlaceholderNode["placeholderType"]> = {
    ctrTitle: "centerTitle",
    subTitle: "subtitle",
    dt: "date",
    ftr: "footer",
    sldNum: "slideNumber",
    pic: "picture",
    tbl: "table",
    obj: "object",
  };
  return value ? (map[value] ?? (value as PlaceholderNode["placeholderType"])) : "body";
}
function importThemes(document: RichDocument<PresentationDocumentRoot>, pkg: OpcPackage) {
  for (const part of pkg.parts.values())
    if (part.contentType.includes("theme") && part.xml) {
      const dom = parseXml(part.xml);
      const theme = dom.documentElement;
      const name = attr(theme, "name");
      document.themes.themes[part.uri] = {
        id: part.uri,
        ...(name ? { name } : {}),
        colors: Object.fromEntries(
          elementsByLocalName(dom, "clrScheme").flatMap((scheme) =>
            childElements(scheme).map((color) => {
              const value = color.firstChild ? attr(color.firstChild as Element, "val") : undefined;
              return [
                color.localName,
                color.firstChild
                  ? { type: color.firstChild.nodeName, ...(value ? { value } : {}) }
                  : {},
              ];
            }),
          ),
        ),
        fonts: {},
        formats: { nativeXml: part.xml },
      };
    }
}

import { createRichDocument, type RichDocument } from "../model/core/document.js";
import type { BaseNode } from "../model/core/node.js";
import type { ParagraphNode, RunProperties, TextRunNode } from "../model/text/index.js";
import type {
  FlowBlockNode,
  FlowDocumentRoot,
  HeaderFooterPart,
  SectionNode,
  TableCellNode,
  TableNode,
  TableRowNode,
} from "../model/word/index.js";
import { stableNodeId } from "../id.js";
import { loadOpcPackage, packageToNativeStore, type OpcPackage } from "../ooxml/opc/package.js";
import { attr, childElements, elementPath, elementsByLocalName, parseXml } from "../ooxml/xml.js";
import { mainPartUri, relatedPartUri } from "./common.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const makeId = (partUri: string, element: Element, prefix: string) => {
  const nativeId = attr(element, "id");
  return stableNodeId(
    {
      sourceFormat: "docx",
      partUri,
      xmlPath: elementPath(element),
      ...(nativeId ? { nativeId } : {}),
    },
    prefix,
  );
};
const source = (partUri: string, element: Element) => ({
  sourceFormat: "docx" as const,
  partUri,
  xmlPath: elementPath(element),
});

export function importDocx(input: Uint8Array): RichDocument<FlowDocumentRoot> {
  const pkg = loadOpcPackage(input);
  const documentUri = mainPartUri(pkg);
  const part = pkg.parts.get(documentUri);
  if (!part?.xml) throw new Error(`DOCX main document is missing: ${documentUri}`);
  const dom = parseXml(part.xml);
  const body = elementsByLocalName(dom, "body")[0];
  if (!body) throw new Error("DOCX has no w:body");
  const sections: SectionNode[] = [];
  let blocks: FlowBlockNode[] = [];
  for (const element of childElements(body)) {
    if (element.localName === "p") blocks.push(parseParagraph(element, documentUri));
    else if (element.localName === "tbl") blocks.push(parseTable(element, documentUri));
    else if (element.localName === "sdt") blocks.push(parseSdt(element, documentUri));
    else if (element.localName === "sectPr") {
      sections.push(parseSection(element, documentUri, blocks));
      blocks = [];
    } else blocks.push(unknown(element, documentUri));
  }
  if (blocks.length || !sections.length)
    sections.push({
      id: `${makeId(documentUri, body, "section")}_implicit`,
      type: "section",
      blocks,
      pageProperties: {},
    });
  const headers = parseHeaderFooters(pkg, documentUri, "header");
  const footers = parseHeaderFooters(pkg, documentUri, "footer");
  const root: FlowDocumentRoot = {
    id: makeId(documentUri, dom.documentElement, "flow"),
    type: "flowDocument",
    sections,
    ...(headers.length ? { headers } : {}),
    ...(footers.length ? { footers } : {}),
    source: source(documentUri, dom.documentElement),
  };
  const result = createRichDocument(`docx:${root.id}`, root);
  result.nativeStore = packageToNativeStore(pkg);
  importStylesAndNumbering(result, pkg, documentUri);
  return result;
}

function parseParagraph(element: Element, partUri: string): ParagraphNode {
  const properties: ParagraphNode["properties"] = {};
  const pPr = childElements(element, "pPr")[0];
  if (pPr) {
    const style = childElements(pPr, "pStyle")[0];
    const align = childElements(pPr, "jc")[0];
    const styleId = style && attr(style, "val");
    const alignment = align && (attr(align, "val") as ParagraphNode["properties"]["alignment"]);
    if (styleId) properties.styleId = styleId;
    if (alignment) properties.alignment = alignment;
  }
  const runs: ParagraphNode["runs"] = [];
  for (const runElement of childElements(element, "r")) {
    const runProperties = parseRunProperties(childElements(runElement, "rPr")[0]);
    for (const child of childElements(runElement)) {
      if (child.localName === "t")
        runs.push({
          id: makeId(partUri, child, "text"),
          type: "textRun",
          text: child.textContent ?? "",
          properties: runProperties,
          source: source(partUri, child),
        });
      else if (child.localName === "tab")
        runs.push({
          id: makeId(partUri, child, "tab"),
          type: "tab",
          source: source(partUri, child),
        });
      else if (child.localName === "br")
        runs.push({
          id: makeId(partUri, child, "break"),
          type: "break",
          kind: (attr(child, "type") as "page" | "column") ?? "line",
          source: source(partUri, child),
        });
      else if (child.localName === "drawing") {
        const blip = elementsByLocalName(child, "blip")[0];
        const resourceId = blip && attr(blip, "embed");
        if (resourceId)
          runs.push({
            id: makeId(partUri, child, "image"),
            type: "inlineImage",
            resourceId,
            source: source(partUri, child),
          });
      }
    }
  }
  return {
    id: makeId(partUri, element, "p"),
    type: "paragraph",
    runs,
    properties,
    source: source(partUri, element),
  };
}
function parseRunProperties(element?: Element): RunProperties {
  if (!element) return {};
  const result: RunProperties = {};
  if (childElements(element, "b").length) result.bold = true;
  if (childElements(element, "i").length) result.italic = true;
  const size = childElements(element, "sz")[0];
  if (size) result.fontSize = { value: Number(attr(size, "val") ?? 0) / 2, unit: "pt" };
  const color = childElements(element, "color")[0];
  const value = color && attr(color, "val");
  if (value && value !== "auto") result.color = { type: "rgb", value };
  return result;
}
function parseTable(element: Element, partUri: string): TableNode {
  const rows: TableRowNode[] = childElements(element, "tr").map((row) => ({
    id: makeId(partUri, row, "tr"),
    type: "tableRow",
    cells: childElements(row, "tc").map((cell) => parseCell(cell, partUri)),
    source: source(partUri, row),
  }));
  const grid = childElements(childElements(element, "tblGrid")[0] ?? element, "gridCol").map(
    (column) => ({ width: { value: Number(attr(column, "w") ?? 0), unit: "twip" as const } }),
  );
  return {
    id: makeId(partUri, element, "tbl"),
    type: "table",
    rows,
    grid: { columns: grid },
    properties: {},
    source: source(partUri, element),
  };
}
function parseCell(element: Element, partUri: string): TableCellNode {
  const blocks = childElements(element)
    .filter(
      (child) => child.localName === "p" || child.localName === "tbl" || child.localName === "sdt",
    )
    .map((child) =>
      child.localName === "p"
        ? parseParagraph(child, partUri)
        : child.localName === "tbl"
          ? parseTable(child, partUri)
          : parseSdt(child, partUri),
    );
  const gridSpan = elementsByLocalName(element, "gridSpan")[0];
  return {
    id: makeId(partUri, element, "tc"),
    type: "tableCell",
    blocks,
    ...(gridSpan ? { columnSpan: Number(attr(gridSpan, "val") ?? 1) } : {}),
    source: source(partUri, element),
  };
}
function parseSdt(
  element: Element,
  partUri: string,
): import("../model/word/index.js").StructuredDocumentTagNode {
  const content = childElements(element, "sdtContent")[0];
  const tagElement = elementsByLocalName(childElements(element, "sdtPr")[0] ?? element, "tag")[0];
  const tag = tagElement && attr(tagElement, "val");
  const blocks = content
    ? childElements(content)
        .filter((child) => ["p", "tbl", "sdt"].includes(child.localName))
        .map((child) =>
          child.localName === "p"
            ? parseParagraph(child, partUri)
            : child.localName === "tbl"
              ? parseTable(child, partUri)
              : parseSdt(child, partUri),
        )
    : [];
  return {
    id: makeId(partUri, element, "sdt"),
    type: "structuredDocumentTag",
    ...(tag ? { tag } : {}),
    blocks,
    source: source(partUri, element),
  };
}
function parseSection(element: Element, partUri: string, blocks: FlowBlockNode[]): SectionNode {
  const size = childElements(element, "pgSz")[0];
  const margin = childElements(element, "pgMar")[0];
  return {
    id: makeId(partUri, element, "section"),
    type: "section",
    blocks,
    pageProperties: {
      ...(size
        ? {
            width: { value: Number(attr(size, "w") ?? 0), unit: "twip" },
            height: { value: Number(attr(size, "h") ?? 0), unit: "twip" },
          }
        : {}),
      ...(margin
        ? {
            margins: Object.fromEntries(
              ["top", "right", "bottom", "left"].map((key) => [
                key,
                { value: Number(attr(margin, key) ?? 0), unit: "twip" },
              ]),
            ),
          }
        : {}),
    },
    headers: childElements(element, "headerReference").map((x) => attr(x, "id") ?? ""),
    footers: childElements(element, "footerReference").map((x) => attr(x, "id") ?? ""),
    source: source(partUri, element),
  };
}
function parseHeaderFooters(
  pkg: OpcPackage,
  documentUri: string,
  kind: "header" | "footer",
): HeaderFooterPart[] {
  return (pkg.parts.get(documentUri)?.relationships ?? [])
    .filter((rel) => rel.type.endsWith(`/${kind}`))
    .flatMap((rel) => {
      const uri = relatedPartUri(pkg, documentUri, rel.id);
      const xml = uri && pkg.parts.get(uri)?.xml;
      if (!uri || !xml) return [];
      const dom = parseXml(xml);
      return [
        {
          id: makeId(uri, dom.documentElement, kind),
          type: kind,
          partUri: uri,
          blocks: childElements(dom.documentElement)
            .filter((x) => x.localName === "p" || x.localName === "tbl")
            .map((x) => (x.localName === "p" ? parseParagraph(x, uri) : parseTable(x, uri))),
          source: source(uri, dom.documentElement),
        },
      ];
    });
}
function importStylesAndNumbering(
  document: RichDocument<FlowDocumentRoot>,
  pkg: OpcPackage,
  documentUri: string,
) {
  for (const rel of pkg.parts.get(documentUri)?.relationships ?? []) {
    const uri = relatedPartUri(pkg, documentUri, rel.id);
    const xml = uri && pkg.parts.get(uri)?.xml;
    if (!uri || !xml) continue;
    if (rel.type.endsWith("/styles"))
      for (const style of elementsByLocalName(parseXml(xml), "style")) {
        const id = attr(style, "styleId");
        const name = attr(elementsByLocalName(style, "name")[0] ?? style, "val");
        const basedOn = attr(elementsByLocalName(style, "basedOn")[0] ?? style, "val");
        const runProperties = parseRunProperties(childElements(style, "rPr")[0]);
        const paragraphProperties = childElements(style, "pPr")[0];
        const alignment =
          paragraphProperties &&
          attr(childElements(paragraphProperties, "jc")[0] ?? paragraphProperties, "val");
        if (id)
          document.styles.styles[id] = {
            id,
            kind:
              (attr(style, "type") as "paragraph" | "character" | "table" | "numbering") ??
              "paragraph",
            ...(name ? { name } : {}),
            ...(basedOn ? { basedOn } : {}),
            properties: { ...runProperties, ...(alignment ? { alignment } : {}) },
            native: { partUri: uri, xmlPath: elementPath(style) },
          };
      }
    if (rel.type.endsWith("/numbering"))
      document.metadata.custom = { ...document.metadata.custom, numberingPartUri: uri };
  }
}
function unknown(element: Element, partUri: string): BaseNode<"unknown"> {
  return {
    id: makeId(partUri, element, "unknown"),
    type: "unknown",
    source: source(partUri, element),
    native: {
      qualifiedName: element.nodeName,
      namespace: element.namespaceURI ?? W_NS,
      rawXml: element.toString(),
    },
  };
}

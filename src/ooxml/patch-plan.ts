import { findByElementPath, parseXml, serializeXml } from "./xml.js";
import type { OpcPackage } from "./opc/package.js";

export type XmlPatch =
  | { op: "setText"; partUri: string; path: string; text: string; preserveSpace?: boolean }
  | { op: "remove"; partUri: string; path: string }
  | {
      op: "insertBefore" | "insertAfter" | "appendChild";
      partUri: string;
      path: string;
      xml: string;
    }
  | {
      op: "setAttributes";
      partUri: string;
      path: string;
      attributes: Record<string, string | null>;
    };

export interface XmlPatchPlan {
  patches: XmlPatch[];
}
export const createPatchPlan = (): XmlPatchPlan => ({ patches: [] });
export const addPatch = (plan: XmlPatchPlan, patch: XmlPatch): void => {
  plan.patches.push(patch);
};

export function applyXmlPatchPlan(pkg: OpcPackage, plan?: XmlPatchPlan): void {
  if (!plan?.patches.length) return;
  const grouped = new Map<string, XmlPatch[]>();
  for (const patch of plan.patches) {
    const list = grouped.get(patch.partUri) ?? [];
    list.push(patch);
    grouped.set(patch.partUri, list);
  }
  for (const [uri, patches] of grouped) {
    const part = pkg.parts.get(uri);
    if (!part?.xml) throw new Error(`Patch target is not an XML part: ${uri}`);
    const document = parseXml(part.xml);
    for (const patch of patches) {
      const target = findByElementPath(document, patch.path);
      if (!target) {
        if (patch.op === "remove") continue;
        throw new Error(`XML patch target not found: ${uri}${patch.path}`);
      }
      if (patch.op === "setText") {
        while (target.firstChild) target.removeChild(target.firstChild);
        target.appendChild(document.createTextNode(patch.text));
        if (patch.preserveSpace ?? /^\s|\s$/.test(patch.text))
          target.setAttribute("xml:space", "preserve");
      } else if (patch.op === "remove") target.parentNode?.removeChild(target);
      else if (patch.op === "setAttributes")
        for (const [name, value] of Object.entries(patch.attributes))
          value === null ? target.removeAttribute(name) : target.setAttribute(name, value);
      else {
        const inserted = parseFragment(patch.xml);
        const node = inserted.cloneNode(true);
        if (patch.op === "appendChild") target.appendChild(node);
        else if (patch.op === "insertBefore") target.parentNode?.insertBefore(node, target);
        else target.parentNode?.insertBefore(node, target.nextSibling);
      }
    }
    part.xml = serializeXml(document);
  }
}

function parseFragment(xml: string): Element {
  const wrapper = parseXml(
    `<rdt:root xmlns:rdt="urn:rdt" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</rdt:root>`,
  );
  for (let node = wrapper.documentElement.firstChild; node; node = node.nextSibling)
    if (node.nodeType === 1) return node as Element;
  throw new Error("XML patch fragment contains no element");
}

export const parentElementPath = (path: string, levels = 1): string => {
  let result = path;
  for (let i = 0; i < levels; i++) result = result.slice(0, result.lastIndexOf("/"));
  return result;
};

import type { RichDocument } from "../model/core/document.js";
import type { GraphicNode } from "../model/drawing/index.js";
import type { PresentationDocumentRoot, SlideElementNode } from "../model/presentation/index.js";
import type { TextBody } from "../model/text/index.js";
import { nativeStoreToPackage, writeOpcPackage } from "../ooxml/opc/package.js";
import { findByElementPath, parseXml, serializeXml } from "../ooxml/xml.js";
import { applyXmlPatchPlan } from "../ooxml/patch-plan.js";

export function exportPptx(document: RichDocument<PresentationDocumentRoot>): Uint8Array {
  if (!document.nativeStore) throw new Error("PPTX export requires the preserved NativeStore");
  const pkg = nativeStoreToPackage(document.nativeStore); const updates = new Map<string, Array<{ path: string; text: string }>>();
  const collectText = (body?: TextBody) => body?.paragraphs.forEach((paragraph) => paragraph.runs.forEach((run) => { if (run.type === "textRun" && run.source?.xmlPath) { const list = updates.get(run.source.partUri) ?? []; list.push({ path: run.source.xmlPath, text: run.text }); updates.set(run.source.partUri, list); } }));
  const collectElement = (element: SlideElementNode | GraphicNode) => { if ((element.type === "shape" || element.type === "placeholder") && "textBody" in element) collectText(element.textBody); if (element.type === "group") element.elements.forEach(collectElement); };
  document.root.slideMasters.forEach((x) => x.shapes.forEach(collectElement)); document.root.slideLayouts.forEach((x) => x.shapes.forEach(collectElement)); document.root.slides.forEach((x) => x.shapes.forEach(collectElement));
  for (const [uri, changes] of updates) { const part = pkg.parts.get(uri); if (!part?.xml) continue; const dom = parseXml(part.xml); for (const change of changes) { const element = findByElementPath(dom, change.path); if (!element) throw new Error(`Cannot map RDT text back to ${uri}${change.path}`); while (element.firstChild) element.removeChild(element.firstChild); element.appendChild(dom.createTextNode(change.text)); } part.xml = serializeXml(dom); }
  applyXmlPatchPlan(pkg, document.patchPlan);
  return writeOpcPackage(pkg);
}

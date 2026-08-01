import { randomUUID } from "node:crypto";
import type { RichDocument } from "../../model/core/document.js";
import type { DocumentRoot } from "../../model/core/document.js";
import type { TextRunNode } from "../../model/text/index.js";
import {
  addPatch,
  createPatchPlan,
  parentElementPath,
  type XmlPatch,
} from "../../ooxml/patch-plan.js";
import { attr, elementsByLocalName, findByElementPath, parseXml } from "../../ooxml/xml.js";
import { indexRdtNodes } from "../markers/common.js";
import { parseMarker } from "../markers/marker.js";
import type { TemplateBindingMap } from "../bindings/binding-map.js";
import { validateBindingMap, flatten } from "../bindings/binding-validator.js";
import type { ConstraintWarning } from "../model/constraints.js";
import type { LlmTemplateDocument } from "../model/template-document.js";
import type { TemplateOperation, TemplateOperationBatch } from "../model/operations.js";
import type {
  CollectionNode,
  FieldGroupNode,
  ListNode,
  TemplateField,
  TemplateNode,
  TextNode,
} from "../model/template-node.js";

export interface ExecutionResult {
  applied: number;
  warnings: ConstraintWarning[];
  createdItemIds: string[];
}
type InsertAfterPatch = { op: "insertAfter"; partUri: string; path: string; xml: string };
interface RuntimeCollectionItem {
  patch: InsertAfterPatch;
  values: Record<string, string | string[]>;
  prototypeXml: string;
}
const runtimeItems = new WeakMap<LlmTemplateDocument, Map<string, RuntimeCollectionItem>>();

export function executeTemplateOperations(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  batch: TemplateOperationBatch | TemplateOperation[],
): ExecutionResult {
  const operations = Array.isArray(batch) ? batch : batch.operations;
  if (!Array.isArray(batch)) {
    if (batch.documentId !== template.documentId)
      throw new Error("Operation documentId does not match template");
    if (batch.revision !== template.revision) throw new Error("Operation revision is stale");
  }
  const bindingIssues = validateBindingMap(template, bindings, document);
  if (bindingIssues.length)
    throw new Error(
      `Binding validation failed: ${bindingIssues.map((item) => item.message).join("; ")}`,
    );
  const warnings: ConstraintWarning[] = [];
  const createdItemIds: string[] = [];
  for (const operation of operations) {
    const created = validateAndExecute(document, template, bindings, operation, warnings);
    if (created) createdItemIds.push(created);
  }
  return { applied: operations.length, warnings, createdItemIds };
}

function validateAndExecute(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  bindings: TemplateBindingMap,
  operation: TemplateOperation,
  warnings: ConstraintWarning[],
): string | undefined {
  const target = flatten(template.root).find((node) => node.id === operation.targetId);
  if (!target) throw new Error(`Unknown operation target: ${operation.targetId}`);
  if (!target.editable) throw new Error(`Target is not editable: ${target.id}`);
  if (
    target.constraints?.hard?.allowedOperations &&
    !target.constraints.hard.allowedOperations.includes(operation.op)
  )
    throw new Error(`Operation ${operation.op} is not allowed on ${target.id}`);
  const characterCount =
    operation.op === "setText"
      ? operation.value.length
      : operation.op === "setList"
        ? operation.items.join("").length
        : operation.op === "appendListItem"
          ? operation.value.length
          : undefined;
  if (
    characterCount !== undefined &&
    target.constraints?.hard?.maxCharacters !== undefined &&
    characterCount > target.constraints.hard.maxCharacters
  )
    throw new Error(`Hard maxCharacters exceeded for ${target.id}`);
  if (
    characterCount !== undefined &&
    target.constraints?.soft?.recommendedCharacters !== undefined &&
    characterCount > target.constraints.soft.recommendedCharacters
  )
    warnings.push({
      code: "recommended-characters-exceeded",
      targetId: target.id,
      message: `${target.id} exceeds recommended character count`,
    });
  if (
    operation.op === "setText" &&
    target.constraints?.soft?.maxLines !== undefined &&
    operation.value.split(/\r?\n/).length > target.constraints.soft.maxLines
  )
    warnings.push({
      code: "max-lines-exceeded",
      targetId: target.id,
      message: `${target.id} may overflow its recommended line count`,
    });
  const itemCount =
    operation.op === "setList"
      ? operation.items.length
      : operation.op === "appendListItem" && target.type === "list"
        ? target.items.length + 1
        : operation.op === "appendCollectionItem" && target.type === "collection"
          ? target.items.length + 1
          : undefined;
  if (
    itemCount !== undefined &&
    target.constraints?.soft?.recommendedItems !== undefined &&
    itemCount > target.constraints.soft.recommendedItems
  )
    warnings.push({
      code: "recommended-items-exceeded",
      targetId: target.id,
      message: `${target.id} exceeds recommended item count`,
    });
  if (operation.op === "setText") executeSetText(document, target, bindings, operation.value);
  else if (operation.op === "setList")
    executeSetList(document, requireList(target), bindings, operation.items);
  else if (operation.op === "appendListItem") {
    const list = requireList(target);
    executeSetList(document, list, bindings, [
      ...list.items.map((item) => item.value),
      operation.value,
    ]);
  } else if (operation.op === "removeListItem") {
    const list = requireList(target);
    const index = operation.itemId
      ? list.items.findIndex((item) => item.id === operation.itemId)
      : (operation.index ?? -1);
    if (index < 0 || index >= list.items.length)
      throw new Error(`List item not found in ${list.id}`);
    const values = list.items.map((item) => item.value);
    values.splice(index, 1);
    executeSetList(document, list, bindings, values);
  } else if (operation.op === "appendCollectionItem")
    return appendCollectionItem(
      document,
      template,
      requireCollection(target),
      bindings,
      operation.value,
    );
  else if (operation.op === "updateCollectionItem")
    updateCollectionItem(
      document,
      template,
      requireCollection(target),
      operation.itemId,
      operation.value,
      bindings,
    );
  else
    removeCollectionItem(document, template, requireCollection(target), operation.itemId, bindings);
}

function executeSetText(
  document: RichDocument<DocumentRoot>,
  target: TemplateNode,
  bindings: TemplateBindingMap,
  value: string,
) {
  if (target.type !== "text") throw new Error(`setText requires text target: ${target.id}`);
  const binding = bindings.bindings[target.id];
  if (!binding) throw new Error(`Missing binding: ${target.id}`);
  if (binding.writeStrategy.type === "insertShapeText") {
    insertPptxTextShape(document, binding, [value], false);
    (target as TextNode).value = value;
    return;
  }
  const nodes = indexRdtNodes(document).byId;
  const runs = binding.sourceNodeIds
    .map((id) => nodes.get(id))
    .filter((node): node is TextRunNode => node?.type === "textRun" && "text" in node);
  if (!runs.length) throw new Error(`Binding has no text runs: ${target.id}`);
  runs[0]!.text = value;
  for (const run of runs.slice(1)) run.text = "";
  (target as TextNode).value = value;
}
function executeSetList(
  document: RichDocument<DocumentRoot>,
  list: ListNode,
  bindings: TemplateBindingMap,
  values: string[],
) {
  const max = list.constraints?.hard?.maxItems;
  if (max !== undefined && values.length > max)
    throw new Error(`Hard maxItems exceeded for ${list.id}`);
  const binding = bindings.bindings[list.id];
  if (!binding) throw new Error(`Missing binding: ${list.id}`);
  if (binding.writeStrategy.type === "insertShapeText") {
    if (values.length) insertPptxTextShape(document, binding, values, true);
    list.items = values.map((value, index) => ({ id: `${list.id}:item:${index + 1}`, value }));
    return;
  }
  const nodes = indexRdtNodes(document).byId;
  const runs = binding.sourceNodeIds
    .map((id) => nodes.get(id))
    .filter((node): node is TextRunNode => node?.type === "textRun" && "text" in node);
  if (!values.length && binding.writeStrategy.allowEmpty) {
    for (const run of runs) run.text = "";
    list.items = [];
    return;
  }
  if (!values.length) throw new Error(`List ${list.id} requires at least one item`);
  if (!runs.length) throw new Error(`List binding has no text runs: ${list.id}`);
  values.slice(0, runs.length).forEach((value, i) => {
    runs[i]!.text = value;
  });
  const plan = (document.patchPlan ??= createPatchPlan());
  for (const run of runs.slice(values.length)) {
    if (run.source?.xmlPath)
      addPatch(plan, {
        op: "remove",
        partUri: run.source.partUri,
        path: parentElementPath(run.source.xmlPath, 2),
      });
  }
  if (values.length > runs.length) {
    const anchor = runs.at(-1)!;
    if (!anchor.source?.xmlPath || !document.nativeStore)
      throw new Error(`List ${list.id} cannot clone an unmapped paragraph`);
    const paragraphPath = parentElementPath(anchor.source.xmlPath, 2);
    const part = document.nativeStore.parts[anchor.source.partUri];
    if (!part?.xml) throw new Error(`List source part is missing: ${anchor.source.partUri}`);
    const paragraph = findByElementPath(parseXml(part.xml), paragraphPath);
    if (!paragraph) throw new Error(`List paragraph source is missing: ${paragraphPath}`);
    for (const value of values.slice(runs.length).reverse()) {
      const clone = paragraph.cloneNode(true) as Element;
      const texts = elementsByLocalName(clone, "t");
      if (!texts.length) throw new Error(`List paragraph has no text node: ${paragraphPath}`);
      texts[0]!.textContent = value;
      for (const text of texts.slice(1)) text.textContent = "";
      addPatch(plan, {
        op: "insertAfter",
        partUri: anchor.source.partUri,
        path: paragraphPath,
        xml: clone.toString(),
      });
    }
  }
  list.items = values.map((value, index) => ({
    id: list.items[index]?.id ?? `${list.id}:item:${randomUUID()}`,
    value,
  }));
}

function insertPptxTextShape(
  document: RichDocument<DocumentRoot>,
  binding: NonNullable<TemplateBindingMap["bindings"][string]>,
  values: string[],
  bullets: boolean,
) {
  const strategy = binding.writeStrategy;
  if (strategy.type !== "insertShapeText") throw new Error("Expected insertShapeText strategy");
  const location = binding.locations?.[0];
  if (document.format !== "pptx" || !location?.partUri || !location.xmlPath)
    throw new Error(`Insert-shape binding is incomplete: ${binding.templateNodeId}`);
  const nativeXml = document.nativeStore?.parts[location.partUri]?.xml;
  const nativeId = nextPptxNonVisualId(nativeXml);
  const paragraphs = values
    .map(
      (value) =>
        `<a:p>${bullets ? '<a:pPr marL="285750" indent="-142875"><a:buChar char="•"/></a:pPr>' : ""}<a:r><a:rPr lang="en-US" sz="900"/><a:t>${escapeXml(value)}</a:t></a:r><a:endParaRPr lang="en-US" sz="900"/></a:p>`,
    )
    .join("");
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${nativeId}" name="RDT ${escapeXml(binding.templateNodeId)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(strategy.x)}" y="${Math.round(strategy.y)}"/><a:ext cx="${Math.round(strategy.width)}" cy="${Math.round(strategy.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
  addPatch((document.patchPlan ??= createPatchPlan()), {
    op: "insertAfter",
    partUri: location.partUri,
    path: location.xmlPath,
    xml,
  });
}
function nextPptxNonVisualId(xml?: string) {
  return (
    Math.max(
      1,
      ...[...(xml ?? "").matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => Number(match[1])),
    ) + 1
  );
}
function escapeXml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function appendCollectionItem(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  collection: CollectionNode,
  bindings: TemplateBindingMap,
  values: Record<string, string | string[]>,
): string {
  const max = collection.maxItems ?? collection.constraints?.hard?.maxItems;
  if (max !== undefined && collection.items.length >= max)
    throw new Error(`Hard maxItems exceeded for ${collection.id}`);
  const prototype = bindings.prototypes[collection.prototypeId];
  const location = prototype?.locations?.[0];
  if (!prototype?.nativeXml || !location?.partUri || !location.xmlPath)
    throw new Error(`Collection prototype is incomplete: ${collection.id}`);
  const itemId = `${collection.id}.item.${randomUUID()}`;
  const xml = instantiatePrototype(prototype.nativeXml, values, itemId, document.format);
  const patch: InsertAfterPatch = {
    op: "insertAfter",
    partUri: location.partUri,
    path: location.xmlPath,
    xml,
  };
  const plan = (document.patchPlan ??= createPatchPlan());
  addPatch(plan, patch);
  const fields = fieldsFromValues(itemId, values);
  collection.items.push({ id: itemId, type: "fieldGroup", label: itemId, editable: false, fields });
  const map = runtimeItems.get(template) ?? new Map();
  map.set(itemId, { patch, values, prototypeXml: prototype.nativeXml });
  runtimeItems.set(template, map);
  return itemId;
}
function updateCollectionItem(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  collection: CollectionNode,
  itemId: string,
  values: Record<string, string | string[]>,
  bindings?: TemplateBindingMap,
) {
  const item = collection.items.find((candidate) => candidate.id === itemId);
  const runtime = runtimeItems.get(template)?.get(itemId);
  if (!item) throw new Error(`Collection item not found: ${itemId}`);
  if (runtime) {
    runtime.values = { ...runtime.values, ...values };
    runtime.patch.xml = instantiatePrototype(
      runtime.prototypeXml,
      runtime.values,
      itemId,
      document.format,
    );
    item.fields = fieldsFromValues(itemId, runtime.values);
    return;
  }
  const nodes = indexRdtNodes(document).byId;
  for (const [key, value] of Object.entries(values)) {
    const binding = bindings?.bindings[`${itemId}.${key}`];
    const runs =
      binding?.sourceNodeIds
        .map((id) => nodes.get(id))
        .filter((node): node is TextRunNode => node?.type === "textRun" && "text" in node) ?? [];
    if (!runs.length) throw new Error(`Collection field binding not found: ${itemId}.${key}`);
    runs[0]!.text = valueText(value);
    for (const run of runs.slice(1)) run.text = "";
  }
  item.fields = fieldsFromValues(
    itemId,
    Object.fromEntries(
      item.fields
        .map((field) => [field.key, field.items?.map((x) => x.value) ?? field.value ?? ""])
        .concat(Object.entries(values)),
    ),
  );
}
function removeCollectionItem(
  document: RichDocument<DocumentRoot>,
  template: LlmTemplateDocument,
  collection: CollectionNode,
  itemId: string,
  bindings?: TemplateBindingMap,
) {
  if (collection.minItems !== undefined && collection.items.length <= collection.minItems)
    throw new Error(`Collection ${collection.id} cannot go below minItems`);
  const index = collection.items.findIndex((item) => item.id === itemId);
  const runtime = runtimeItems.get(template)?.get(itemId);
  if (index < 0) throw new Error(`Collection item not found: ${itemId}`);
  if (runtime) {
    const patches = document.patchPlan?.patches;
    const patchIndex = patches?.indexOf(runtime.patch) ?? -1;
    if (patchIndex >= 0) patches!.splice(patchIndex, 1);
    runtimeItems.get(template)?.delete(itemId);
  } else {
    const location = bindings?.bindings[itemId]?.locations?.[0];
    if (!location?.partUri || !location.xmlPath)
      throw new Error(`Collection item root binding not found: ${itemId}`);
    addPatch((document.patchPlan ??= createPatchPlan()), {
      op: "remove",
      partUri: location.partUri,
      path: location.xmlPath,
    });
  }
  collection.items.splice(index, 1);
}
function instantiatePrototype(
  nativeXml: string,
  values: Record<string, string | string[]>,
  itemId: string,
  format: "docx" | "pptx",
): string {
  const wrapper = parseXml(
    `<rdt:root xmlns:rdt="urn:rdt" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${nativeXml}</rdt:root>`,
  );
  const root = [
    ...Array.from({ length: wrapper.documentElement.childNodes.length }, (_, i) =>
      wrapper.documentElement.childNodes.item(i),
    ),
  ].find((node) => node?.nodeType === 1) as Element | undefined;
  if (!root) throw new Error("Prototype XML has no root element");
  if (format === "docx") {
    for (const sdt of [root, ...elementsByLocalName(root, "sdt")].filter(
      (element) => element.localName === "sdt",
    )) {
      const tag = elementsByLocalName(sdt, "tag")[0];
      const raw = tag ? attr(tag, "val") : undefined;
      const marker = raw ? parseMarker(raw) : undefined;
      if (marker?.kind === "field") {
        setElementTexts(sdt, valueText(values[marker.id]));
        tag!.setAttribute("w:val", `rdt:field:${itemId}.${marker.id}`);
      } else if (marker?.kind === "prototype") tag!.setAttribute("w:val", `rdt:group:${itemId}`);
    }
  } else {
    let nextId = 100000 + Math.floor(Math.random() * 800000);
    for (const properties of elementsByLocalName(root, "cNvPr")) {
      const marker =
        parseMarker(attr(properties, "name") ?? "") ?? parseMarker(attr(properties, "descr") ?? "");
      if (marker?.kind === "field") {
        let shape = properties.parentNode?.parentNode as Element;
        while (shape && !["sp", "pic", "grpSp"].includes(shape.localName))
          shape = shape.parentNode as Element;
        if (shape) setElementTexts(shape, valueText(values[marker.id]));
        properties.setAttribute("name", `rdt:field:${itemId}.${marker.id}`);
      }
      properties.setAttribute("id", String(nextId++));
    }
  }
  return root.toString();
}
function setElementTexts(element: Element, value: string) {
  const texts = elementsByLocalName(element, "t");
  if (!texts.length) return;
  texts[0]!.textContent = value;
  for (const text of texts.slice(1)) text.textContent = "";
}
function valueText(value: string | string[] | undefined) {
  return Array.isArray(value) ? value.join("\n") : (value ?? "");
}
function fieldsFromValues(
  itemId: string,
  values: Record<string, string | string[]>,
): TemplateField[] {
  return Object.entries(values).map(([key, value]) => ({
    id: `${itemId}.${key}`,
    key,
    type: Array.isArray(value) ? "list" : "text",
    ...(Array.isArray(value)
      ? { items: value.map((item, index) => ({ id: `${itemId}.${key}:${index}`, value: item })) }
      : { value }),
    editable: true,
  }));
}
function requireList(node: TemplateNode): ListNode {
  if (node.type !== "list") throw new Error(`List operation requires list target: ${node.id}`);
  return node;
}
function requireCollection(node: TemplateNode): CollectionNode {
  if (node.type !== "collection")
    throw new Error(`Collection operation requires collection target: ${node.id}`);
  return node;
}

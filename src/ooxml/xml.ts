import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export const parseXml = (xml: string): Document => {
  const errors: string[] = [];
  const document = new DOMParser({
    onError: (level, message) => {
      if (level === "error" || level === "fatalError") errors.push(message);
    },
  }).parseFromString(xml, "application/xml") as unknown as Document;
  if (errors.length) throw new Error(`Invalid XML: ${errors.join("; ")}`);
  return document;
};

export const serializeXml = (document: Document): string =>
  new XMLSerializer().serializeToString(document as never);
export const elementsByLocalName = (root: Document | Element, name: string): Element[] => {
  const all = root.getElementsByTagName("*");
  const result: Element[] = [];
  for (let index = 0; index < all.length; index++) {
    const element = all.item(index);
    if (element && (element.localName === name || element.nodeName.split(":").at(-1) === name))
      result.push(element);
  }
  return result;
};
export const childElements = (node: Node, localName?: string): Element[] => {
  const result: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling)
    if (
      child.nodeType === 1 &&
      (!localName ||
        (child as Element).localName === localName ||
        child.nodeName.split(":").at(-1) === localName)
    )
      result.push(child as Element);
  return result;
};
export const attr = (element: Element, localName: string): string | undefined => {
  for (let index = 0; index < element.attributes.length; index++) {
    const item = element.attributes.item(index);
    if (item && (item.localName === localName || item.name.split(":").at(-1) === localName))
      return item.value;
  }
  return undefined;
};

export function elementPath(element: Element): string {
  const segments: string[] = [];
  for (
    let current: Element | null = element;
    current;
    current = current.parentNode?.nodeType === 1 ? (current.parentNode as Element) : null
  ) {
    const name = current.localName || current.nodeName.split(":").at(-1) || current.nodeName;
    let index = 1;
    for (let sibling = current.previousSibling; sibling; sibling = sibling.previousSibling)
      if (
        sibling.nodeType === 1 &&
        ((sibling as Element).localName || sibling.nodeName.split(":").at(-1)) === name
      )
        index++;
    segments.unshift(`${name}[${index}]`);
  }
  return `/${segments.join("/")}`;
}

export function findByElementPath(document: Document, path: string): Element | undefined {
  let current: Element | undefined = document.documentElement;
  const segments = path.replace(/^\//, "").split("/");
  const first = parseSegment(segments.shift());
  if (
    !current ||
    !first ||
    (current.localName || current.nodeName.split(":").at(-1)) !== first.name ||
    first.index !== 1
  )
    return undefined;
  for (const segment of segments) {
    const target = parseSegment(segment);
    if (!target) return undefined;
    current = childElements(current, target.name)[target.index - 1];
    if (!current) return undefined;
  }
  return current;
}
function parseSegment(segment?: string): { name: string; index: number } | undefined {
  const match = segment?.match(/^(.+)\[(\d+)]$/);
  return match?.[1] && match[2] ? { name: match[1], index: Number(match[2]) } : undefined;
}

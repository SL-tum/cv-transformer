import type { OpcPackage } from "../ooxml/opc/package.js";
import { resolveRelationshipTarget } from "../ooxml/opc/relationships.js";

export const OFFICE_DOCUMENT_REL = "/officeDocument";
export function mainPartUri(pkg: OpcPackage): string {
  const relationship = pkg.rootRelationships.find((item) =>
    item.type.endsWith(OFFICE_DOCUMENT_REL),
  );
  const uri = relationship && resolveRelationshipTarget(undefined, relationship);
  if (!uri) throw new Error("Package has no officeDocument root relationship");
  return uri;
}
export const relatedPartUri = (
  pkg: OpcPackage,
  sourceUri: string,
  relationshipId: string,
): string | undefined => {
  const relationship = pkg.parts
    .get(sourceUri)
    ?.relationships.find((item) => item.id === relationshipId);
  return relationship ? resolveRelationshipTarget(sourceUri, relationship) : undefined;
};

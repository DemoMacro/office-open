/**
 * Relationship manager for OOXML .rels file management.
 */
import type { Element } from "@office-open/xml";

import { getFirstLevelElements } from "./xml-patch-utils";

const getIdFromRelationshipId = (relationshipId: string): number => {
  const output = parseInt(relationshipId.substring(3), 10);
  return isNaN(output) ? 0 : output;
};

export const getNextRelationshipIndex = (relationships: Element): number => {
  const relationshipElements = getFirstLevelElements(relationships, "Relationships");

  return (
    relationshipElements
      .map((e) => getIdFromRelationshipId(e.attributes?.Id?.toString() ?? ""))
      .reduce((acc, curr) => Math.max(acc, curr), 0) + 1
  );
};

export const appendRelationship = (
  relationships: Element,
  id: number | string,
  type: string,
  target: string,
  targetMode?: string,
): readonly Element[] => {
  const relationshipElements = getFirstLevelElements(relationships, "Relationships");
  relationshipElements.push({
    attributes: {
      Id: `rId${id}`,
      Target: target,
      TargetMode: targetMode,
      Type: type,
    },
    name: "Relationship",
    type: "element",
  });

  return relationshipElements;
};

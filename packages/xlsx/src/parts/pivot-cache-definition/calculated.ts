import type {
  CalculatedItemOptions,
  CalculatedMemberOptions,
  PivotAreaOptions,
  PivotAreaReferenceOptions,
} from "../pivot/pivot-utils";
import type { PivotXmlElementOptions } from "../pivot/pivot-xml";
import {
  findPivotChildren,
  stringifyPivotXmlElement,
  type PivotXmlAttributeValue,
} from "../pivot/pivot-xml";
import { XlsxParseError } from "../pivot/pivot-xml";

function attributesOf(
  node: PivotXmlElementOptions,
): Record<string, PivotXmlAttributeValue | undefined> {
  return node.attributes ?? {};
}

function numberAttr(node: PivotXmlElementOptions, name: string): number | undefined {
  const value = attributesOf(node)[name];
  return value === undefined ? undefined : Number(value);
}

function stringAttr(node: PivotXmlElementOptions, name: string): string | undefined {
  const value = attributesOf(node)[name];
  return value === undefined ? undefined : String(value);
}

function boolAttr(node: PivotXmlElementOptions, name: string): boolean | undefined {
  const value = attributesOf(node)[name];
  return value === undefined ? undefined : value === "1";
}

function setAttribute(
  attributes: Record<string, PivotXmlAttributeValue | undefined>,
  name: string,
  value: PivotXmlAttributeValue | undefined,
): void {
  if (value !== undefined) attributes[name] = value;
}

function parsePivotArea(node: PivotXmlElementOptions): PivotAreaOptions {
  const result: PivotAreaOptions = {};
  const field = numberAttr(node, "field");
  if (field !== undefined) result.field = field;
  const type = stringAttr(node, "type") as PivotAreaOptions["type"];
  if (type !== undefined) result.type = type;
  if (stringAttr(node, "dataOnly") === "0") result.dataOnly = false;
  if (boolAttr(node, "labelOnly")) result.labelOnly = true;
  if (boolAttr(node, "grandRow")) result.grandRow = true;
  if (boolAttr(node, "grandCol")) result.grandCol = true;
  if (boolAttr(node, "cacheIndex")) result.cacheIndex = true;
  if (stringAttr(node, "outline") === "0") result.outline = false;
  const offset = stringAttr(node, "offset");
  if (offset !== undefined) result.offset = offset;
  if (boolAttr(node, "collapsedLevelsAreSubtotals")) result.collapsedLevelsAreSubtotals = true;
  const axis = stringAttr(node, "axis") as PivotAreaOptions["axis"];
  if (axis !== undefined) result.axis = axis;
  const fieldPosition = numberAttr(node, "fieldPosition");
  if (fieldPosition !== undefined) result.fieldPosition = fieldPosition;
  const referencesNode = findPivotChildren(node, "references")[0];
  if (referencesNode) {
    result.references = findPivotChildren(referencesNode, "reference").map(parsePivotAreaReference);
  }
  return result;
}

function parsePivotAreaReference(node: PivotXmlElementOptions): PivotAreaReferenceOptions {
  const result: PivotAreaReferenceOptions = {};
  const field = numberAttr(node, "field");
  if (field !== undefined) result.field = field;
  const count = numberAttr(node, "count");
  if (count !== undefined) result.count = count;
  if (stringAttr(node, "selected") === "0") result.selected = false;
  for (const name of [
    "byPosition",
    "relative",
    "defaultSubtotal",
    "sumSubtotal",
    "avgSubtotal",
    "countASubtotal",
    "maxSubtotal",
    "minSubtotal",
    "countSubtotal",
    "productSubtotal",
    "stdDevPSubtotal",
    "stdDevSubtotal",
    "varPSubtotal",
    "varSubtotal",
  ] as const) {
    if (boolAttr(node, name)) result[name] = true;
  }
  const indexes = findPivotChildren(node, "x")
    .map((child) => numberAttr(child, "v"))
    .filter((value): value is number => value !== undefined);
  if (indexes.length > 0) result.x = indexes;
  return result;
}

function pivotAreaNode(area: PivotAreaOptions): PivotXmlElementOptions {
  const attributes: Record<string, PivotXmlAttributeValue | undefined> = {};
  setAttribute(attributes, "field", area.field);
  setAttribute(attributes, "type", area.type);
  if (area.dataOnly === false) attributes.dataOnly = "0";
  setAttribute(attributes, "labelOnly", area.labelOnly ? "1" : undefined);
  setAttribute(attributes, "grandRow", area.grandRow ? "1" : undefined);
  setAttribute(attributes, "grandCol", area.grandCol ? "1" : undefined);
  setAttribute(attributes, "cacheIndex", area.cacheIndex ? "1" : undefined);
  if (area.outline === false) attributes.outline = "0";
  setAttribute(attributes, "offset", area.offset);
  setAttribute(
    attributes,
    "collapsedLevelsAreSubtotals",
    area.collapsedLevelsAreSubtotals ? "1" : undefined,
  );
  setAttribute(attributes, "axis", area.axis);
  setAttribute(attributes, "fieldPosition", area.fieldPosition);
  const children = area.references
    ? [
        {
          name: "references",
          attributes: { count: area.references.length },
          children: area.references.map(pivotAreaReferenceNode),
        },
      ]
    : undefined;
  return { name: "pivotArea", attributes, ...(children ? { children } : {}) };
}

function pivotAreaReferenceNode(reference: PivotAreaReferenceOptions): PivotXmlElementOptions {
  const attributes: Record<string, PivotXmlAttributeValue | undefined> = {};
  setAttribute(attributes, "field", reference.field);
  setAttribute(attributes, "count", reference.count);
  if (reference.selected === false) attributes.selected = "0";
  for (const name of [
    "byPosition",
    "relative",
    "defaultSubtotal",
    "sumSubtotal",
    "avgSubtotal",
    "countASubtotal",
    "maxSubtotal",
    "minSubtotal",
    "countSubtotal",
    "productSubtotal",
    "stdDevPSubtotal",
    "stdDevSubtotal",
    "varPSubtotal",
    "varSubtotal",
  ] as const) {
    setAttribute(attributes, name, reference[name] ? "1" : undefined);
  }
  return {
    name: "reference",
    attributes,
    ...(reference.x
      ? { children: reference.x.map((v) => ({ name: "x", attributes: { v } })) }
      : {}),
  };
}

export function calculatedItemNode(item: CalculatedItemOptions): PivotXmlElementOptions {
  const attributes: Record<string, PivotXmlAttributeValue | undefined> = {};
  setAttribute(attributes, "field", item.field);
  setAttribute(attributes, "formula", item.formula);
  return {
    name: "calculatedItem",
    attributes,
    ...(item.pivotArea ? { children: [pivotAreaNode(item.pivotArea)] } : {}),
  };
}

export function calculatedMemberNode(member: CalculatedMemberOptions): PivotXmlElementOptions {
  const attributes: Record<string, PivotXmlAttributeValue | undefined> = {
    name: member.name,
    mdx: member.mdx,
  };
  setAttribute(attributes, "memberName", member.memberName);
  setAttribute(attributes, "hierarchy", member.hierarchy);
  setAttribute(attributes, "parent", member.parent);
  setAttribute(attributes, "solveOrder", member.solveOrder);
  setAttribute(attributes, "set", member.set ? "1" : undefined);
  return { name: "calculatedMember", attributes };
}

export function parseCalculatedItems(
  part: string,
  parentPath: string,
  node: PivotXmlElementOptions,
): CalculatedItemOptions[] {
  return findPivotChildren(node, "calculatedItem").map((item) => {
    if (findPivotChildren(item, "extLst").length > 0) {
      throw new XlsxParseError(
        part,
        `${parentPath}/calculatedItem`,
        "extLst",
        "unsupported calculated item extension",
      );
    }
    const field = numberAttr(item, "field");
    const result: CalculatedItemOptions = {};
    if (field !== undefined) result.field = field;
    const formula = stringAttr(item, "formula");
    if (formula !== undefined) result.formula = formula;
    const pivotArea = findPivotChildren(item, "pivotArea")[0];
    if (pivotArea) result.pivotArea = parsePivotArea(pivotArea);
    return result;
  });
}

export function parseCalculatedMembers(
  part: string,
  parentPath: string,
  node: PivotXmlElementOptions,
): CalculatedMemberOptions[] {
  return findPivotChildren(node, "calculatedMember").map((member) => {
    if (findPivotChildren(member, "extLst").length > 0) {
      throw new XlsxParseError(
        part,
        `${parentPath}/calculatedMember`,
        "extLst",
        "unsupported calculated member extension",
      );
    }
    const name = stringAttr(member, "name");
    const mdx = stringAttr(member, "mdx");
    if (!name || !mdx) {
      throw new XlsxParseError(
        part,
        `${parentPath}/calculatedMember`,
        member.name,
        "name and mdx are required",
      );
    }
    const result: CalculatedMemberOptions = { name, mdx };
    const memberName = stringAttr(member, "memberName");
    if (memberName !== undefined) result.memberName = memberName;
    const hierarchy = stringAttr(member, "hierarchy");
    if (hierarchy !== undefined) result.hierarchy = hierarchy;
    const parent = stringAttr(member, "parent");
    if (parent !== undefined) result.parent = parent;
    const solveOrder = numberAttr(member, "solveOrder");
    if (solveOrder !== undefined) result.solveOrder = solveOrder;
    if (boolAttr(member, "set")) result.set = true;
    return result;
  });
}

export function calculatedItemsXml(items: CalculatedItemOptions[]): string {
  return `<calculatedItems count="${items.length}">${items
    .map((item) => stringifyPivotXmlElement(calculatedItemNode(item)))
    .join("")}</calculatedItems>`;
}

export function calculatedMembersXml(members: CalculatedMemberOptions[]): string {
  return `<calculatedMembers count="${members.length}">${members
    .map((member) => stringifyPivotXmlElement(calculatedMemberNode(member)))
    .join("")}</calculatedMembers>`;
}

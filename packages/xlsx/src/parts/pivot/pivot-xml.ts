import { escapeXml, type Element } from "@office-open/xml";

import { PIVOT_ALLOWED_CHILDREN } from "./pivot-xml-schema";

export type PivotXmlAttributeValue = string | number | boolean;

export interface PivotXmlElementOptions {
  name: string;
  attributes?: Record<string, PivotXmlAttributeValue | undefined>;
  children?: PivotXmlElementOptions[];
  text?: string;
}

export class XlsxParseError extends Error {
  constructor(
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
  ) {
    super(`${part}${path}: ${name}: ${reason}`);
  }
}

export function stringifyPivotXmlElement(node: PivotXmlElementOptions): string {
  const attrs = Object.entries(node.attributes ?? {})
    .filter(([, value]) => value !== undefined)
    .map(([name, value]) => `${name}="${escapeXml(String(value))}"`);
  const open = `<${node.name}${attrs.length ? ` ${attrs.join(" ")}` : ""}`;
  if (!node.children?.length && node.text === undefined) return `${open}/>`;
  const children = (node.children ?? []).map(stringifyPivotXmlElement).join("");
  return `${open}>${escapeXml(node.text ?? "")}${children}</${node.name}>`;
}

export function parsePivotXmlElement(
  part: string,
  el: Element,
  parentPath = "",
): PivotXmlElementOptions {
  const name = el.name ?? "";
  const path = `${parentPath}/${name}`;
  const children: PivotXmlElementOptions[] = [];
  for (const child of el.elements ?? []) {
    const childName = child.name ?? "";
    if (!childName) continue;
    if (child.instruction || child.comment) continue;
    const allowed = PIVOT_ALLOWED_CHILDREN[name] ?? [];
    if (allowed[0] !== "*" && !allowed.includes(childName)) {
      throw new XlsxParseError(
        part,
        path,
        childName,
        `unexpected child of ${name}; expected one of: ${allowed.join(", ") || "none"}`,
      );
    }
    children.push(parsePivotXmlElement(part, child, path));
  }
  return {
    name,
    ...(el.attributes ? { attributes: { ...el.attributes } } : {}),
    ...(children.length ? { children } : {}),
    ...(el.text !== undefined && el.text !== "" ? { text: String(el.text) } : {}),
  };
}

export function findPivotChildren(
  node: PivotXmlElementOptions,
  name: string,
): PivotXmlElementOptions[] {
  return (node.children ?? []).filter((child) => child.name === name);
}

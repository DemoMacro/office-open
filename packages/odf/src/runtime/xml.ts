import { escapeXml, parse, stringifyElement } from "@office-open/xml";
import type { Element } from "@office-open/xml";

export { stringifyElement };

export type XmlAttributes = Record<string, string | number | boolean | undefined>;

export function escapeText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function xmlElement(
  name: string,
  attributes?: XmlAttributes,
  children: string[] = [],
): string {
  const renderedAttributes = Object.entries(attributes ?? {})
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => ` ${key}="${escapeXml(String(value))}"`)
    .join("");
  return children.length
    ? `<${name}${renderedAttributes}>${children.join("")}</${name}>`
    : `<${name}${renderedAttributes}/>`;
}

/** Parses one XML fragment and returns its first element. */
export function parseXmlElement(xml: string): Element | undefined {
  return parse(xml).elements?.find((element) => element.type === "element");
}

export function childrenNamed(element: Element | undefined, name: string): Element[] {
  return element?.elements?.filter((child) => child.name === name) ?? [];
}

export function childNamed(element: Element | undefined, name: string): Element | undefined {
  return element?.elements?.find((child) => child.name === name);
}

export function textOf(element: Element | undefined): string {
  if (!element) return "";
  if (typeof element.text === "string") return element.text;
  return (element.elements ?? [])
    .filter((child) => child.type === "text")
    .map((child) => String(child.text ?? ""))
    .join("");
}

export function descendantElements(element: Element | undefined): Element[] {
  const result: Element[] = [];
  const pending = [...(element?.elements ?? [])].reverse();
  while (pending.length > 0) {
    const child = pending.pop()!;
    if (child.type !== "element") continue;
    result.push(child);
    for (const grandchild of [...(child.elements ?? [])].reverse()) {
      pending.push(grandchild);
    }
  }
  return result;
}

export function spreadsheetColumnName(column: number): string {
  let value = column;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export function spreadsheetColumnNumber(column: string): number {
  let number = 0;
  for (let index = 0; index < column.length; index += 1) {
    number = number * 26 + column.charCodeAt(index) - 64;
  }
  return number;
}

export function isExternalOdfUrl(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("//") || value.startsWith("../");
}

export function unknownAttributeNames(
  element: Element,
  allowedAttributes: readonly string[],
): string[] {
  return Object.keys(element.attributes ?? {}).filter((name) => !allowedAttributes.includes(name));
}

export function unknownChildElement(
  element: Element,
  allowedChildren: readonly string[],
): Element | undefined {
  return (element.elements ?? []).find(
    (child) => child.type === "element" && !allowedChildren.includes(child.name ?? ""),
  );
}

export function attributeString(element: Element | undefined, name: string): string | undefined {
  if (!element) return undefined;
  const value = element.attributes?.[name];
  return value === undefined ? undefined : String(value);
}

export function attributeNumber(element: Element | undefined, name: string): number | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function lengthToEmu(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const match = /^(-?\d+(?:\.\d+)?)(cm|mm|pt|pc|in|px)$/.exec(value);
  if (!match?.[2]) return undefined;
  const amount = Number(match[1]);
  const unitFactors: Record<string, number> = {
    cm: 360000,
    mm: 36000,
    pt: 12700,
    pc: 152400,
    in: 914400,
    px: 9525,
  };
  const factor = unitFactors[match[2]];
  return factor === undefined ? undefined : Math.round(amount * factor);
}

/** Converts an authored number (EMU) or ODF universal length to EMU. */
export function odfLengthToEmu(value: number | string): number {
  if (typeof value === "number") return value;
  const emu = lengthToEmu(value);
  if (emu === undefined) throw new Error(`Invalid ODF length: ${value}`);
  return emu;
}

export function emuToLength(emu: number | string | undefined, unit: "cm" | "px" = "cm"): string {
  const value = typeof emu === "string" ? lengthToEmu(emu) : emu;
  if (value === undefined) return "0cm";
  if (unit === "px") return `${Number((value / 9525).toFixed(9))}px`;
  const amount = Number((value / 360000).toFixed(6));
  return `${amount}cm`;
}

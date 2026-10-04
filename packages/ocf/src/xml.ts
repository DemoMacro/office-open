import { escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

export type XmlAttributes = Record<string, string | number | boolean | undefined>;

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
  return Math.round(amount * unitFactors[match[2] as "cm"]!);
}

export function emuToLength(emu: number | string | undefined, unit: "cm" | "px" = "cm"): string {
  const value = typeof emu === "string" ? lengthToEmu(emu) : emu;
  if (value === undefined) return "0cm";
  if (unit === "px") return `${value}px`;
  const amount = Number((value / 360000).toFixed(6));
  return `${amount}cm`;
}

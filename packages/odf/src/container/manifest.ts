import { parse, type Element } from "@office-open/xml";

import { OcfManifestError } from "../errors";
import { textOf, xmlElement } from "../runtime/xml";

export const MANIFEST_ELEMENTS = [
  "manifest:manifest",
  "manifest:file-entry",
  "manifest:encryption-data",
  "manifest:algorithm",
  "manifest:key-derivation",
  "manifest:start-key-generation",
  "manifest:keyinfo",
  "manifest:encrypted-key",
  "manifest:encryption-method",
  "manifest:PGPData",
  "manifest:PGPKeyID",
  "manifest:PGPKeyPacket",
  "manifest:CipherData",
  "manifest:CipherValue",
] as const;

export const MANIFEST_ATTRIBUTES = [
  "manifest:algorithm-name",
  "manifest:checksum",
  "manifest:checksum-type",
  "manifest:full-path",
  "manifest:initialisation-vector",
  "manifest:iteration-count",
  "manifest:key-derivation-name",
  "manifest:key-size",
  "manifest:media-type",
  "manifest:PGPAlgorithm",
  "manifest:preferred-view-mode",
  "manifest:salt",
  "manifest:size",
  "manifest:start-key-generation-name",
  "manifest:version",
] as const;

export type ManifestElementName = (typeof MANIFEST_ELEMENTS)[number];
export type ManifestAttributeName = (typeof MANIFEST_ATTRIBUTES)[number];
export type ManifestAttributes = Partial<Record<ManifestAttributeName, string | number | boolean>>;

export interface OcfManifestElement {
  name: ManifestElementName;
  attributes?: ManifestAttributes;
  children?: OcfManifestElement[];
  foreignContentXml?: string;
  text?: string;
}

export interface OcfManifestOptions {
  root: OcfManifestElement;
}

export const MANIFEST_SCHEMA_ELEMENTS = MANIFEST_ELEMENTS.map((name) => ({
  name,
  attributes: MANIFEST_ATTRIBUTES,
}));

const ELEMENT_SET = new Set<string>(MANIFEST_ELEMENTS);
const ATTRIBUTE_SET = new Set<string>(MANIFEST_ATTRIBUTES);

export function parseManifestOptions(xml: string): OcfManifestOptions {
  const root = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true }).elements?.find(
    (element) => element.type === "element",
  );
  if (root?.name !== "manifest:manifest") throw new OcfManifestError("Invalid ODF manifest");
  return { root: parseManifestElement(root, "/manifest:manifest") };
}

function parseManifestElement(element: Element, path: string): OcfManifestElement {
  if (!ELEMENT_SET.has(element.name ?? ""))
    throw new OcfManifestError(`Unknown manifest element: ${path}/${element.name}`);
  const attributes: ManifestAttributes = {};
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (!ATTRIBUTE_SET.has(name))
      throw new OcfManifestError(`Unknown manifest attribute: ${path}/${name}`);
    attributes[name as ManifestAttributeName] = value;
  }
  const foreign =
    element.name === "manifest:algorithm" || element.name === "manifest:key-derivation";
  return {
    name: element.name as ManifestElementName,
    attributes,
    foreignContentXml: foreign
      ? (element.elements ?? [])
          .filter((child): child is Element => child.type === "element")
          .map(foreignXml)
          .join("")
      : undefined,
    text: textOf(element) || undefined,
    children: foreign
      ? undefined
      : (element.elements ?? [])
          .filter((child): child is Element => child.type === "element")
          .map((child) => parseManifestElement(child, `${path}/${element.name}`)),
  };
}

export function manifestOptionsXml(options: OcfManifestOptions): string {
  return manifestElementXml(options.root);
}

function manifestElementXml(element: OcfManifestElement): string {
  return xmlElement(element.name, element.attributes, [
    ...(element.text ? [element.text] : []),
    ...(element.foreignContentXml ? [element.foreignContentXml] : []),
    ...(element.children ?? []).map(manifestElementXml),
  ]);
}

function foreignXml(element: Element): string {
  return xmlElement(element.name ?? "", element.attributes, [
    ...(textOf(element) ? [textOf(element)!] : []),
    ...(element.elements ?? [])
      .filter((child): child is Element => child.type === "element")
      .map(foreignXml),
  ]);
}

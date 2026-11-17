import { parse, type Element } from "@office-open/xml";

import { OcfManifestError } from "../errors";
import { escapeText, textOf, xmlElement } from "../runtime/xml";

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

export type ManifestAttributeName = (typeof MANIFEST_ATTRIBUTES)[number];

const MANIFEST_ELEMENT_ATTRIBUTES: Record<string, readonly string[]> = {
  "manifest:manifest": ["manifest:version"],
  "manifest:file-entry": [
    "manifest:full-path",
    "manifest:version",
    "manifest:media-type",
    "manifest:preferred-view-mode",
    "manifest:size",
  ],
  "manifest:encryption-data": ["manifest:checksum-type", "manifest:checksum"],
  "manifest:algorithm": ["manifest:algorithm-name", "manifest:initialisation-vector"],
  "manifest:key-derivation": [
    "manifest:key-derivation-name",
    "manifest:salt",
    "manifest:iteration-count",
    "manifest:key-size",
  ],
  "manifest:start-key-generation": ["manifest:start-key-generation-name", "manifest:key-size"],
  "manifest:encryption-method": ["manifest:PGPAlgorithm"],
};

export interface OcfEncryptedKeyOptions {
  algorithm?: string;
  keyId?: string;
  keyPacket?: string;
  cipherValue?: string;
}

export interface OcfAlgorithmOptions {
  name: string;
  initialisationVector: string;
  foreignContentXml?: string;
}

export interface OcfKeyDerivationOptions {
  name: string;
  salt?: string;
  iterationCount?: number;
  keySize?: number;
}

export interface OcfStartKeyGenerationOptions {
  name: string;
  keySize?: number;
}

export interface OcfEncryptionDataOptions {
  checksumType: string;
  checksum: string;
  algorithm: OcfAlgorithmOptions;
  startKeyGeneration?: OcfStartKeyGenerationOptions;
  keyDerivation: OcfKeyDerivationOptions;
}

export interface OcfManifestFileEntryOptions {
  fullPath: string;
  version?: string;
  mediaType?: string;
  preferredViewMode?: string;
  /** Size in bytes. */
  size?: number;
  encryptionData?: OcfEncryptionDataOptions;
}

export interface OcfManifestOptions {
  /** Manifest grammar version (`1.3`). */
  version: string;
  /** File entries in package-declared order. */
  entries: OcfManifestFileEntryOptions[];
  /** Encrypted package keys, in source order. */
  encryptedKeys?: OcfEncryptedKeyOptions[];
}

export const MANIFEST_SCHEMA_ELEMENTS = MANIFEST_ELEMENTS.map((name) => ({
  name,
  attributes: MANIFEST_ATTRIBUTES,
}));

export function parseManifestOptions(xml: string): OcfManifestOptions {
  const document = parse(xml, {
    ignoreDeclaration: true,
    ignoreDoctype: true,
    normalizeNamespaces: {
      "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0": "manifest",
      "urn:oasis:names:tc:opendocument:xmlns:manifest:1.3": "manifest",
    },
  });
  const elements = document.elements ?? [];
  const roots = elements.filter((element) => element.type === "element");
  if (
    roots.length !== 1 ||
    elements.some(
      (element) =>
        element.type !== "element" &&
        element.type !== "comment" &&
        element.type !== "instruction" &&
        !(element.type === "text" && String(element.text ?? "").trim().length === 0),
    )
  ) {
    throw new OcfManifestError("Invalid ODF manifest: expected exactly one root element");
  }
  const root = roots[0];
  if (root?.name !== "manifest:manifest") throw new OcfManifestError("Invalid ODF manifest");
  rejectUnknownManifestAttributes(root);
  return parseManifestModel(root);
}

function rejectUnknownManifestAttributes(element: Element, allowForeignAttributes = false): void {
  const name = element.name ?? "";
  const allowed = MANIFEST_ELEMENT_ATTRIBUTES[name] ?? [];
  for (const attribute of Object.keys(element.attributes ?? {})) {
    if (
      attribute === "xmlns" ||
      attribute.startsWith("xmlns:") ||
      allowed.includes(attribute) ||
      (allowForeignAttributes && !name.startsWith("manifest:"))
    )
      continue;
    throw new OcfManifestError(`Unknown manifest attribute @${attribute}: ${name}`);
  }
  for (const child of element.elements ?? []) {
    if (child.type === "element")
      rejectUnknownManifestAttributes(
        child,
        allowForeignAttributes || name === "manifest:algorithm",
      );
  }
}

function parseManifestModel(root: Element): OcfManifestOptions {
  const namespaces = namespaceDeclarations(root.attributes);
  const children = (root.elements ?? []).filter(
    (element) => element.type === "element",
  ) as Element[];
  if (
    (root.elements ?? []).some(
      (element) => element.type === "text" && String(element.text ?? "").trim().length > 0,
    )
  ) {
    throw new OcfManifestError("Invalid ODF manifest: unexpected text in manifest root");
  }
  const encryptedKeys: OcfEncryptedKeyOptions[] = [];
  const entries: OcfManifestFileEntryOptions[] = [];
  let seenFileEntry = false;
  for (const child of children) {
    if (child.name === "manifest:file-entry") {
      seenFileEntry = true;
      entries.push(parseFileEntry(child, namespaces));
    } else if (child.name === "manifest:encrypted-key") {
      if (seenFileEntry) throw unknownManifestChild(child);
      encryptedKeys.push(parseEncryptedKey(child));
    } else {
      throw unknownManifestChild(child);
    }
  }
  if (entries.length === 0) throw new OcfManifestError("Manifest has no file entries");
  return {
    version:
      root.attributes?.["manifest:version"] === undefined
        ? "1.3"
        : requiredManifestAttribute(root, "version", "/manifest:manifest"),
    entries,
    ...(encryptedKeys.length > 0 ? { encryptedKeys } : {}),
  };
}

function parseFileEntry(
  element: Element,
  namespaces: Record<string, string | number | undefined>,
): OcfManifestFileEntryOptions {
  const fileEntryNamespaces = namespaceDeclarations(element.attributes, namespaces);
  const path = "/manifest:manifest/manifest:file-entry";
  for (const child of element.elements ?? []) {
    if (child.type === "element" && child.name !== "manifest:encryption-data")
      throw unknownManifestChild(child);
  }
  const size = requiredOrUndefined(element, "size", path);
  const encryptions = (element.elements ?? []).filter(
    (child): child is Element =>
      child.type === "element" && child.name === "manifest:encryption-data",
  );
  if (encryptions.length > 1)
    throw new OcfManifestError("Duplicate manifest encryption-data: " + path);
  const encryption = encryptions[0];
  return {
    fullPath: requiredManifestAttribute(element, "full-path", path),
    ...(requiredOrUndefined(element, "version", path)
      ? { version: requiredManifestAttribute(element, "version", path) }
      : {}),
    ...(requiredOrUndefined(element, "media-type", path)
      ? { mediaType: requiredManifestAttribute(element, "media-type", path) }
      : {}),
    ...(requiredOrUndefined(element, "preferred-view-mode", path)
      ? {
          preferredViewMode: requiredManifestAttribute(element, "preferred-view-mode", path),
        }
      : {}),
    ...(size === undefined ? {} : { size: numberAttribute(size, `${path}/@manifest:size`) }),
    ...(encryption ? { encryptionData: parseEncryptionData(encryption, fileEntryNamespaces) } : {}),
  };
}

function namespaceDeclarations(
  attributes: Element["attributes"],
  inherited: Record<string, string | number | undefined> = {},
): Record<string, string | number | undefined> {
  const declarations = { ...inherited };
  for (const [name, value] of Object.entries(attributes ?? {})) {
    if (name === "xmlns" || name.startsWith("xmlns:")) declarations[name] = value;
  }
  return declarations;
}

function parseEncryptionData(
  element: Element,
  namespaces: Record<string, string | number | undefined>,
): OcfEncryptionDataOptions {
  const encryptionNamespaces = namespaceDeclarations(element.attributes, namespaces);
  const path = "/manifest:manifest/manifest:file-entry/manifest:encryption-data";
  const children = (element.elements ?? []).filter(
    (child) => child.type === "element",
  ) as Element[];
  const algorithms = children.filter((child) => child.name === "manifest:algorithm");
  const keyDerivations = children.filter((child) => child.name === "manifest:key-derivation");
  const startKeys = children.filter((child) => child.name === "manifest:start-key-generation");
  const expectedNames = [
    "manifest:algorithm",
    ...(startKeys.length === 0 ? [] : ["manifest:start-key-generation"]),
    "manifest:key-derivation",
  ];
  if (
    algorithms.length !== 1 ||
    keyDerivations.length !== 1 ||
    startKeys.length > 1 ||
    children.length !== expectedNames.length ||
    children.some((child, index) => child.name !== expectedNames[index])
  )
    throw new OcfManifestError("Invalid manifest encryption-data: " + path);
  const algorithm = algorithms[0]!;
  return {
    checksumType: requiredManifestAttribute(element, "checksum-type", path),
    checksum: requiredManifestAttribute(element, "checksum", path),
    algorithm: parseAlgorithm(algorithm, encryptionNamespaces),
    ...(startKeys[0] ? { startKeyGeneration: parseStartKeyGeneration(startKeys[0]) } : {}),
    keyDerivation: parseKeyDerivation(keyDerivations[0]!),
  };
}

function parseAlgorithm(
  element: Element,
  namespaces: Record<string, string | number | undefined>,
): OcfAlgorithmOptions {
  const algorithmNamespaces = namespaceDeclarations(element.attributes, namespaces);
  const path = "/manifest:manifest/manifest:file-entry/manifest:encryption-data/manifest:algorithm";
  return {
    name: requiredManifestAttribute(element, "algorithm-name", path),
    initialisationVector: requiredManifestAttribute(element, "initialisation-vector", path),
    ...(element.elements?.length
      ? {
          foreignContentXml: (element.elements ?? [])
            .filter((child): child is Element => child.type === "element")
            .map((child) => foreignXml(child, algorithmNamespaces))
            .join(""),
        }
      : {}),
  };
}

function parseStartKeyGeneration(element: Element): OcfStartKeyGenerationOptions {
  const path =
    "/manifest:manifest/manifest:file-entry/manifest:encryption-data/manifest:start-key-generation";
  return {
    name: requiredManifestAttribute(element, "start-key-generation-name", path),
    ...(requiredOrUndefined(element, "key-size", path)
      ? { keySize: numberAttribute(requiredOrUndefined(element, "key-size", path)!, path) }
      : {}),
  };
}

function parseKeyDerivation(element: Element): OcfKeyDerivationOptions {
  const path =
    "/manifest:manifest/manifest:file-entry/manifest:encryption-data/manifest:key-derivation";
  if ((element.elements ?? []).some((child) => child.type === "element"))
    throw new OcfManifestError("Invalid manifest key-derivation content: " + path);
  return {
    name: requiredManifestAttribute(element, "key-derivation-name", path),
    ...(requiredOrUndefined(element, "salt", path)
      ? { salt: requiredManifestAttribute(element, "salt", path) }
      : {}),
    ...(requiredOrUndefined(element, "iteration-count", path)
      ? {
          iterationCount: numberAttribute(
            requiredOrUndefined(element, "iteration-count", path)!,
            path,
          ),
        }
      : {}),
    ...(requiredOrUndefined(element, "key-size", path)
      ? { keySize: numberAttribute(requiredOrUndefined(element, "key-size", path)!, path) }
      : {}),
  };
}

function parseEncryptedKey(element: Element): OcfEncryptedKeyOptions {
  const path = "/manifest:manifest/manifest:encrypted-key";
  const children = childElements(element);
  const method = optionalChild(children, "manifest:encryption-method", path);
  const keyInfo = requiredChild(children, "manifest:keyinfo", path);
  const keyInfoChildren = childElements(keyInfo);
  const pgpData = requiredChild(keyInfoChildren, "manifest:PGPData", path);
  const pgpChildren = childElements(pgpData);
  const keyId = requiredChild(pgpChildren, "manifest:PGPKeyID", path);
  const keyPacket = optionalChild(pgpChildren, "manifest:PGPKeyPacket", path);
  const cipherData = requiredChild(children, "manifest:CipherData", path);
  const cipherValue = requiredChild(childElements(cipherData), "manifest:CipherValue", path);
  if (
    children.length !== (method ? 3 : 2) ||
    keyInfoChildren.length !== 1 ||
    pgpChildren.length !== (keyPacket ? 2 : 1) ||
    childElements(cipherData).length !== 1
  )
    throw new OcfManifestError(`Invalid manifest encrypted-key structure: ${path}`);
  return {
    ...(method ? { algorithm: requiredManifestAttribute(method, "PGPAlgorithm", path) } : {}),
    keyId: textOf(keyId),
    ...(keyPacket ? { keyPacket: textOf(keyPacket) } : {}),
    cipherValue: textOf(cipherValue),
  };
}

function childElements(element: Element): Element[] {
  return (element.elements ?? []).filter((child): child is Element => child.type === "element");
}

function requiredChild(children: readonly Element[], name: string, path: string): Element {
  const matches = children.filter((child) => child.name === name);
  if (matches.length !== 1)
    throw new OcfManifestError(`Invalid manifest encrypted-key structure: ${path}`);
  return matches[0]!;
}

function optionalChild(
  children: readonly Element[],
  name: string,
  path: string,
): Element | undefined {
  const matches = children.filter((child) => child.name === name);
  if (matches.length > 1)
    throw new OcfManifestError(`Invalid manifest encrypted-key structure: ${path}`);
  return matches[0];
}

function requiredManifestAttribute(element: Element, name: string, path: string): string {
  const value = element.attributes?.[`manifest:${name}`];
  if (value === undefined)
    throw new OcfManifestError(`Manifest is missing @manifest:${name}: ${path}`);
  return String(value);
}

function requiredOrUndefined(element: Element, name: string, _path: string): string | undefined {
  const value = element.attributes?.[`manifest:${name}`];
  return value === undefined ? undefined : String(value);
}

function numberAttribute(value: string, path: string): number {
  const result = Number(value);
  if (!Number.isInteger(result) || result < 0)
    throw new OcfManifestError(`Invalid non-negative integer: ${path}`);
  return result;
}

export function manifestOptionsXml(options: OcfManifestOptions): string {
  return manifestModelXml(options);
}

function manifestModelXml(options: OcfManifestOptions): string {
  return xmlElement(
    "manifest:manifest",
    {
      "xmlns:manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
      "manifest:version": options.version,
    },
    [
      ...(options.encryptedKeys ?? []).map(encryptedKeyXml),
      ...options.entries.map(manifestFileEntryXml),
    ],
  );
}

function encryptedKeyXml(key: OcfEncryptedKeyOptions): string {
  return xmlElement("manifest:encrypted-key", undefined, [
    ...(key.algorithm
      ? [xmlElement("manifest:encryption-method", { "manifest:PGPAlgorithm": key.algorithm })]
      : []),
    xmlElement("manifest:keyinfo", undefined, [
      xmlElement("manifest:PGPData", undefined, [
        ...(key.keyId ? [xmlElement("manifest:PGPKeyID", undefined, [key.keyId])] : []),
        ...(key.keyPacket ? [xmlElement("manifest:PGPKeyPacket", undefined, [key.keyPacket])] : []),
      ]),
    ]),
    xmlElement("manifest:CipherData", undefined, [
      xmlElement("manifest:CipherValue", undefined, [key.cipherValue ?? ""]),
    ]),
  ]);
}

export function manifestFileEntryXml(entry: OcfManifestFileEntryOptions): string {
  return xmlElement(
    "manifest:file-entry",
    {
      "manifest:full-path": entry.fullPath,
      "manifest:version": entry.version,
      "manifest:media-type": entry.mediaType,
      "manifest:preferred-view-mode": entry.preferredViewMode,
      "manifest:size": entry.size,
    },
    entry.encryptionData ? [encryptionDataXml(entry.encryptionData)] : [],
  );
}

function encryptionDataXml(data: OcfEncryptionDataOptions): string {
  return xmlElement(
    "manifest:encryption-data",
    {
      "manifest:checksum-type": data.checksumType,
      "manifest:checksum": data.checksum,
    },
    [
      xmlElement(
        "manifest:algorithm",
        {
          "manifest:algorithm-name": data.algorithm.name,
          "manifest:initialisation-vector": data.algorithm.initialisationVector,
        },
        [data.algorithm.foreignContentXml ?? ""],
      ),
      ...(data.startKeyGeneration
        ? [
            xmlElement("manifest:start-key-generation", {
              "manifest:start-key-generation-name": data.startKeyGeneration.name,
              "manifest:key-size": data.startKeyGeneration.keySize,
            }),
          ]
        : []),
      xmlElement("manifest:key-derivation", {
        "manifest:key-derivation-name": data.keyDerivation.name,
        "manifest:salt": data.keyDerivation.salt,
        "manifest:iteration-count": data.keyDerivation.iterationCount,
        "manifest:key-size": data.keyDerivation.keySize,
      }),
    ],
  );
}

function foreignXml(
  element: Element,
  inheritedNamespaces: Record<string, string | number | undefined> = {},
): string {
  const attributes = { ...element.attributes };
  for (const [name, value] of Object.entries(inheritedNamespaces)) {
    if (attributes[name] === undefined) attributes[name] = value;
  }
  const children = (element.elements ?? []).flatMap((child) => {
    if (child.type === "element") return [foreignXml(child)];
    if (child.type === "text") return [escapeText(String(child.text ?? ""))];
    if (child.type === "cdata") {
      const value = String(child.cdata ?? "").replaceAll("]]>", "]]]]><![CDATA[>");
      return [`<![CDATA[${value}]]>`];
    }
    return [];
  });
  if (element.text !== undefined) children.unshift(escapeText(String(element.text)));
  if (element.cdata !== undefined) {
    children.push(`<![CDATA[${String(element.cdata).replaceAll("]]>", "]]]]><![CDATA[>")}]]>`);
  }
  return xmlElement(element.name ?? "", attributes, children);
}

function unknownManifestChild(element: Element): OcfManifestError {
  return new OcfManifestError("Unknown manifest element: " + (element.name ?? ""));
}

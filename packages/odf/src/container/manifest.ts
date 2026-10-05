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

export type ManifestAttributeName = (typeof MANIFEST_ATTRIBUTES)[number];

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
  const root = parse(xml, {
    ignoreDeclaration: true,
    ignoreDoctype: true,
    normalizeNamespaces: {
      "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0": "manifest",
      "urn:oasis:names:tc:opendocument:xmlns:manifest:1.3": "manifest",
    },
  }).elements?.find((element) => element.type === "element");
  if (root?.name !== "manifest:manifest") throw new OcfManifestError("Invalid ODF manifest");
  return parseManifestModel(root);
}

function parseManifestModel(root: Element): OcfManifestOptions {
  const children = (root.elements ?? []).filter(
    (element) => element.type === "element",
  ) as Element[];
  const encryptedKeys: OcfEncryptedKeyOptions[] = [];
  const entries: OcfManifestFileEntryOptions[] = [];
  let seenFileEntry = false;
  for (const child of children) {
    if (child.name === "manifest:file-entry") {
      seenFileEntry = true;
      entries.push(parseFileEntry(child));
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

function parseFileEntry(element: Element): OcfManifestFileEntryOptions {
  const path = "/manifest:manifest/manifest:file-entry";
  for (const child of element.elements ?? []) {
    if (child.type === "element" && child.name !== "manifest:encryption-data")
      throw unknownManifestChild(child);
  }
  const size = requiredOrUndefined(element, "size", path);
  const encryption = (element.elements ?? []).find(
    (child) => child.type === "element" && child.name === "manifest:encryption-data",
  ) as Element | undefined;
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
    ...(encryption ? { encryptionData: parseEncryptionData(encryption) } : {}),
  };
}

function parseEncryptionData(element: Element): OcfEncryptionDataOptions {
  const path = "/manifest:manifest/manifest:file-entry/manifest:encryption-data";
  const children = (element.elements ?? []).filter(
    (child) => child.type === "element",
  ) as Element[];
  const algorithm = children.find((child) => child.name === "manifest:algorithm");
  const keyDerivation = children.find((child) => child.name === "manifest:key-derivation");
  const startKey = children.find((child) => child.name === "manifest:start-key-generation");
  if (!algorithm || !keyDerivation)
    throw new OcfManifestError(`Invalid manifest encryption-data: ${path}`);
  return {
    checksumType: requiredManifestAttribute(element, "checksum-type", path),
    checksum: requiredManifestAttribute(element, "checksum", path),
    algorithm: parseAlgorithm(algorithm),
    ...(startKey ? { startKeyGeneration: parseStartKeyGeneration(startKey) } : {}),
    keyDerivation: parseKeyDerivation(keyDerivation),
  };
}

function parseAlgorithm(element: Element): OcfAlgorithmOptions {
  const path = "/manifest:manifest/manifest:file-entry/manifest:encryption-data/manifest:algorithm";
  return {
    name: requiredManifestAttribute(element, "algorithm-name", path),
    initialisationVector: requiredManifestAttribute(element, "initialisation-vector", path),
    ...(element.elements?.length
      ? {
          foreignContentXml: (element.elements ?? [])
            .filter((child): child is Element => child.type === "element")
            .map(foreignXml)
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
  const method = findDescendant(element, "manifest:encryption-method");
  const keyId = findDescendant(element, "manifest:PGPKeyID");
  const keyPacket = findDescendant(element, "manifest:PGPKeyPacket");
  const cipherValue = findDescendant(element, "manifest:CipherValue");
  return {
    ...(method ? { algorithm: requiredManifestAttribute(method, "PGPAlgorithm", path) } : {}),
    ...(keyId ? { keyId: textOf(keyId) } : {}),
    ...(keyPacket ? { keyPacket: textOf(keyPacket) } : {}),
    ...(cipherValue ? { cipherValue: textOf(cipherValue) } : {}),
  };
}

function findDescendant(element: Element, name: string): Element | undefined {
  for (const child of element.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === name) return child;
    const descendant = findDescendant(child, name);
    if (descendant) return descendant;
  }
  return undefined;
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

function foreignXml(element: Element): string {
  return xmlElement(element.name ?? "", element.attributes, [
    ...(textOf(element) ? [textOf(element)!] : []),
    ...(element.elements ?? [])
      .filter((child): child is Element => child.type === "element")
      .map(foreignXml),
  ]);
}

function unknownManifestChild(element: Element): OcfManifestError {
  return new OcfManifestError("Unknown manifest element: " + (element.name ?? ""));
}

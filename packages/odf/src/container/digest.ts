import { DEFAULT_ARCHIVE_READ_LIMITS, ParsedArchive, strFromU8 } from "@office-open/core";
import { parse, type Element } from "@office-open/xml";

import { OcfManifestError } from "../errors";
import {
  parseManifestOptions,
  type OcfEncryptionDataOptions,
  type OcfEncryptedKeyOptions,
} from "./manifest";

export interface OcfEntrySemanticDigest {
  fullPath: string;
  version?: string;
  mediaType?: string;
  preferredViewMode?: string;
  size?: number;
  encryptionData?: OcfEncryptionDataOptions;
  binaryDigest?: string;
  xmlDigest?: string;
}

export interface OcfMemberSemanticDigest {
  path: string;
  binaryDigest?: string;
  xmlDigest?: string;
}

export interface OcfSemanticDigest {
  version: string;
  mimeType?: string;
  entries: OcfEntrySemanticDigest[];
  members: OcfMemberSemanticDigest[];
  encryptedKeys?: OcfEncryptedKeyOptions[];
}

export function ocfSemanticDigest(data: Uint8Array): OcfSemanticDigest {
  let archive: Record<string, Uint8Array>;
  try {
    const parsed = new ParsedArchive(data, undefined, DEFAULT_ARCHIVE_READ_LIMITS);
    archive = Object.fromEntries(
      parsed.keys().map((path) => [path, new Uint8Array(parsed.getRaw(path)!)]),
    );
  } catch (cause) {
    throw new OcfManifestError("Invalid OCF ZIP package", { cause });
  }
  const manifestXml = archive["META-INF/manifest.xml"];
  if (!manifestXml) throw new OcfManifestError("ODF package is missing META-INF/manifest.xml");
  const manifest = parseManifestOptions(strFromU8(manifestXml));
  const members = Object.keys(archive)
    .filter((path) => path !== "mimetype" && path !== "META-INF/manifest.xml")
    .sort()
    .map((path) => {
      const bytes = archive[path]!;
      return path.endsWith(".xml")
        ? { path, xmlDigest: xmlDigest(strFromU8(bytes)) }
        : { path, binaryDigest: binaryDigest(bytes) };
    });
  for (const entry of manifest.entries) {
    if (
      entry.fullPath !== "/" &&
      !entry.fullPath.endsWith("/") &&
      archive[entry.fullPath] === undefined &&
      !isOptionalMissingManifestPath(entry.fullPath)
    ) {
      throw new OcfManifestError(
        `ODF manifest declares missing package member: ${entry.fullPath}`,
        {
          fullPath: entry.fullPath,
        },
      );
    }
  }
  return {
    version: manifest.version,
    ...(archive.mimetype === undefined ? {} : { mimeType: strFromU8(archive.mimetype) }),
    ...(manifest.encryptedKeys === undefined ? {} : { encryptedKeys: manifest.encryptedKeys }),
    members,
    entries: manifest.entries.map((entry) => {
      const bytes = archive[entry.fullPath];
      return {
        fullPath: entry.fullPath,
        ...(entry.version === undefined ? {} : { version: entry.version }),
        ...(entry.mediaType === undefined ? {} : { mediaType: entry.mediaType }),
        ...(entry.preferredViewMode === undefined
          ? {}
          : { preferredViewMode: entry.preferredViewMode }),
        ...(entry.size === undefined ? {} : { size: entry.size }),
        ...(entry.encryptionData === undefined ? {} : { encryptionData: entry.encryptionData }),
        ...(bytes === undefined
          ? {}
          : entry.fullPath.endsWith(".xml")
            ? { xmlDigest: xmlDigest(strFromU8(bytes)) }
            : { binaryDigest: binaryDigest(bytes) }),
      };
    }),
  };
}

export function compareOcfSemantic(left: Uint8Array, right: Uint8Array): string[] {
  const differences: string[] = [];
  const leftDigest = ocfSemanticDigest(left);
  const rightDigest = ocfSemanticDigest(right);
  if (leftDigest.version !== rightDigest.version)
    differences.push(`manifest.version: ${leftDigest.version} != ${rightDigest.version}`);
  if (leftDigest.mimeType !== rightDigest.mimeType)
    differences.push(`mimetype: ${leftDigest.mimeType} != ${rightDigest.mimeType}`);
  if (JSON.stringify(leftDigest.encryptedKeys) !== JSON.stringify(rightDigest.encryptedKeys))
    differences.push("manifest.encryptedKeys differ");
  const memberLength = Math.max(leftDigest.members.length, rightDigest.members.length);
  for (let index = 0; index < memberLength; index += 1) {
    const leftMember = leftDigest.members[index];
    const rightMember = rightDigest.members[index];
    if (!leftMember || !rightMember) {
      differences.push(
        `member[${index}]: ${leftMember?.path ?? "(missing)"} != ${rightMember?.path ?? "(missing)"}`,
      );
      continue;
    }
    if (JSON.stringify(leftMember) !== JSON.stringify(rightMember))
      differences.push(
        `member[${index}]: ${JSON.stringify(leftMember)} != ${JSON.stringify(rightMember)}`,
      );
  }
  const length = Math.max(leftDigest.entries.length, rightDigest.entries.length);
  for (let index = 0; index < length; index += 1) {
    const leftEntry = leftDigest.entries[index];
    const rightEntry = rightDigest.entries[index];
    if (!leftEntry || !rightEntry) {
      differences.push(
        `entry[${index}]: ${leftEntry?.fullPath ?? "(missing)"} != ${rightEntry?.fullPath ?? "(missing)"}`,
      );
      continue;
    }
    for (const key of [
      "fullPath",
      "version",
      "mediaType",
      "preferredViewMode",
      "size",
      "encryptionData",
      "binaryDigest",
      "xmlDigest",
    ] as const) {
      const leftValue = JSON.stringify(leftEntry[key]);
      const rightValue = JSON.stringify(rightEntry[key]);
      if (leftValue !== rightValue)
        differences.push(`entry[${index}].${key}: ${leftValue} != ${rightValue}`);
    }
  }
  return differences;
}

function xmlDigest(xml: string): string {
  const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
  const root = document.elements?.find((element) => element.type === "element");
  return digestText(elementDigest(root));
}

function elementDigest(element: Element | undefined): string {
  if (!element) return "";
  const attributes = Object.entries(element.attributes ?? {})
    .map(([name, value]) => `${name}=${String(value)}`)
    .sort()
    .join(";");
  const children = (element.elements ?? [])
    .filter((child) => child.type !== "comment")
    .map((child) => {
      if (child.type === "element") return elementDigest(child);
      if (child.type === "cdata") return String(child.cdata ?? "");
      return String(child.text ?? "");
    })
    .filter(Boolean)
    .join("|");
  return `<${element.name ?? ""} ${attributes}>${children}`;
}

function binaryDigest(bytes: Uint8Array): string {
  let value = "";
  for (let index = 0; index < bytes.length; index += 4096)
    value += String.fromCharCode(...bytes.subarray(index, index + 4096));
  return digestText(value);
}

function digestText(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

function isOptionalMissingManifestPath(path: string): boolean {
  return (
    path === "Thumbnails/thumbnail.png" ||
    path === "styles.xml" ||
    path === "meta.xml" ||
    path === "settings.xml" ||
    path === "layout-cache" ||
    path.startsWith("Pictures/") ||
    path.startsWith("Oleobject")
  );
}

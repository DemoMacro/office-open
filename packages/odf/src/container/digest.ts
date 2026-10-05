import { parse, type Element } from "@office-open/xml";
import { strFromU8, unzipSync } from "fflate";

import { OcfManifestError } from "../errors";
import { parseManifestOptions } from "./manifest";

export interface OcfEntrySemanticDigest {
  fullPath: string;
  version?: string;
  mediaType?: string;
  preferredViewMode?: string;
  size?: number;
  binaryDigest?: string;
  xmlDigest?: string;
}

export interface OcfSemanticDigest {
  version: string;
  entries: OcfEntrySemanticDigest[];
}

export function ocfSemanticDigest(data: Uint8Array): OcfSemanticDigest {
  let archive: Record<string, Uint8Array>;
  try {
    archive = unzipSync(data);
  } catch (cause) {
    throw new OcfManifestError("Invalid OCF ZIP package", { cause });
  }
  const manifestXml = archive["META-INF/manifest.xml"];
  if (!manifestXml) throw new OcfManifestError("ODF package is missing META-INF/manifest.xml");
  const manifest = parseManifestOptions(strFromU8(manifestXml));
  return {
    version: manifest.version,
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
      "binaryDigest",
      "xmlDigest",
    ] as const) {
      if (leftEntry[key] !== rightEntry[key])
        differences.push(`entry[${index}].${key}: ${leftEntry[key]} != ${rightEntry[key]}`);
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
    .map((child) =>
      child.type === "element" ? elementDigest(child) : String(child.text ?? "").trim(),
    )
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

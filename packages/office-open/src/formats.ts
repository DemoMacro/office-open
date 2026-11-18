import { CompoundFileReader } from "@office-open/core/cfb";
import {
  OOXML_PACKAGE_FORMATS,
  ParsedArchive,
  resolveRelationshipTarget,
} from "@office-open/core/opc";
import type { OoxmlPackageFormat } from "@office-open/core/opc";
import { parse } from "@office-open/xml";

export type OfficeFormat =
  | "doc"
  | "docx"
  | "docm"
  | "dotx"
  | "dotm"
  | "xls"
  | "xlsx"
  | "xlsm"
  | "xltx"
  | "xltm"
  | "ppt"
  | "pptx"
  | "pptm"
  | "potx"
  | "potm"
  | "rtf"
  | "odt"
  | "ods"
  | "odp"
  | "encrypted-ooxml";

export type OfficeFormatFamily = "document" | "workbook" | "presentation" | "unknown";
export type OfficeFormatContainer = "cfb" | "zip" | "text";

export interface OfficeFormatInfo {
  format: OfficeFormat;
  family: OfficeFormatFamily;
  container: OfficeFormatContainer;
}

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;
const ODF_MIME_TYPES: Record<string, OfficeFormat> = {
  "application/vnd.oasis.opendocument.text": "odt",
  "application/vnd.oasis.opendocument.spreadsheet": "ods",
  "application/vnd.oasis.opendocument.presentation": "odp",
};
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const LEGACY_CONTENT_TYPES_NAMESPACE = "http://schemas.microsoft.com/package/2005/06/content-types";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const LEGACY_RELATIONSHIPS_NAMESPACE = "http://schemas.microsoft.com/package/2005/06/relationships";

export function detectOffice(input: Uint8Array | string): OfficeFormatInfo {
  if (typeof input === "string") {
    if (
      input
        .replace(/^\uFEFF\s*/, "")
        .trimStart()
        .startsWith("{\\rtf")
    ) {
      return { format: "rtf", family: "document", container: "text" };
    }
    throw new Error("Unable to detect office format from text input");
  }

  if (input.byteLength >= 8 && CFB_SIGNATURE.every((byte, index) => input[index] === byte)) {
    return detectCfbFormat(input);
  }
  if (input.byteLength >= 4 && input[0] === 0x50 && input[1] === 0x4b) {
    return detectZipFormat(input);
  }

  const text = new TextDecoder("latin1", { fatal: false }).decode(
    input[0] === 0xef && input[1] === 0xbb && input[2] === 0xbf
      ? input.subarray(3, 67)
      : input.subarray(0, 64),
  );
  if (
    text
      .replace(/^\uFEFF\s*/, "")
      .trimStart()
      .startsWith("{\\rtf")
  ) {
    return { format: "rtf", family: "document", container: "text" };
  }
  throw new Error("Unable to detect office format");
}

/**
 * Rewrite the primary part's Override declaration when the caller re-types a
 * parsed package (the requested format differs from the source table);
 * otherwise the preserved source declaration wins the content-type output.
 */
export function withRequestedPackageVariant<
  T extends {
    contentTypes?: { overrides: ReadonlyArray<{ partName: string; contentType: string }> };
  },
>(format: keyof typeof OOXML_PACKAGE_FORMATS, options: T): T {
  const info = OOXML_PACKAGE_FORMATS[format];
  const overrides = options.contentTypes?.overrides;
  if (!overrides) return options;
  const partName = `/${info.mainPartPath}`.toLowerCase();
  const declared = overrides.find(
    (override) => override.partName.toLowerCase() === partName,
  )?.contentType;
  if (declared === undefined || declared === info.mainContentType) return options;
  const isFamilyMain = Object.values(OOXML_PACKAGE_FORMATS).some(
    (candidate) => candidate.family === info.family && candidate.mainContentType === declared,
  );
  if (!isFamilyMain) return options;
  return {
    ...options,
    contentTypes: {
      ...options.contentTypes,
      overrides: overrides.map((override) =>
        override.partName.toLowerCase() === partName
          ? { ...override, contentType: info.mainContentType }
          : override,
      ),
    },
  };
}

function detectCfbFormat(data: Uint8Array): OfficeFormatInfo {
  const reader = new CompoundFileReader(data);
  const paths = new Set(reader.entries.map((entry) => entry.path.toLowerCase()));
  if (paths.has("worddocument")) return { format: "doc", family: "document", container: "cfb" };
  if (paths.has("workbook") || paths.has("book")) {
    return { format: "xls", family: "workbook", container: "cfb" };
  }
  if (paths.has("powerpoint document")) {
    return { format: "ppt", family: "presentation", container: "cfb" };
  }
  if (paths.has("encryptedpackage") && paths.has("encryptioninfo")) {
    return { format: "encrypted-ooxml", family: "unknown", container: "cfb" };
  }
  throw new Error("Unable to determine the CFB office format");
}

function detectZipFormat(data: Uint8Array): OfficeFormatInfo {
  let archive: ParsedArchive;
  try {
    archive = new ParsedArchive(data);
  } catch {
    throw new Error("Unable to read the ZIP office container");
  }

  const mimeType = archive.getRaw("mimetype");
  if (mimeType) {
    const format = ODF_MIME_TYPES[new TextDecoder().decode(mimeType).trim()];
    if (!format) throw new Error("Unsupported ODF mimetype");
    return {
      format,
      family: format === "ods" ? "workbook" : format === "odp" ? "presentation" : "document",
      container: "zip",
    };
  }

  const contentTypes = archive.getRaw("[Content_Types].xml");
  if (!contentTypes) throw new Error("ZIP office container is missing [Content_Types].xml");
  const rootRelationships = archive.getRaw("_rels/.rels");
  if (!rootRelationships) throw new Error("OOXML package is missing root relationships");
  const normalizeNamespaces = {
    [CONTENT_TYPES_NAMESPACE]: "ct",
    [LEGACY_CONTENT_TYPES_NAMESPACE]: "ct",
    [RELATIONSHIPS_NAMESPACE]: "rel",
    [LEGACY_RELATIONSHIPS_NAMESPACE]: "rel",
  };
  const types = parse(new TextDecoder().decode(contentTypes), {
    normalizeNamespaces,
  }).elements?.find((element) => element.type === "element");
  if (types?.name !== "ct:Types") throw new Error("Invalid OOXML content types part");
  const relationships = parse(new TextDecoder().decode(rootRelationships), {
    normalizeNamespaces,
  }).elements?.find((element) => element.type === "element");
  if (relationships?.name !== "rel:Relationships")
    throw new Error("Invalid OOXML root relationships part");
  const mainPartTarget = (relationships.elements ?? []).find((relationship) => {
    const type = relationship.attributes?.Type;
    const targetMode = relationship.attributes?.TargetMode;
    return (
      relationship.name === "rel:Relationship" &&
      typeof type === "string" &&
      type.endsWith("/officeDocument") &&
      (typeof targetMode !== "string" || targetMode.toLowerCase() !== "external") &&
      typeof relationship.attributes?.Target === "string"
    );
  })?.attributes?.Target;
  if (typeof mainPartTarget !== "string")
    throw new Error("OOXML package is missing its officeDocument relationship");
  const mainPartPath = resolveRelationshipTarget("", mainPartTarget.replaceAll("\\", "/"));
  const mainContentType = (types.elements ?? []).find((entry) => {
    const partName = entry.attributes?.PartName;
    return (
      entry.name === "ct:Override" &&
      typeof partName === "string" &&
      partName.toLowerCase() === `/${mainPartPath}`.toLowerCase()
    );
  })?.attributes?.ContentType;
  const format = (Object.keys(OOXML_PACKAGE_FORMATS) as OoxmlPackageFormat[]).find(
    (candidate) => OOXML_PACKAGE_FORMATS[candidate].mainContentType === mainContentType,
  );
  if (!format) throw new Error("Unsupported OOXML package");
  const info = OOXML_PACKAGE_FORMATS[format];
  return {
    format,
    family:
      info.family === "wordprocessing"
        ? "document"
        : info.family === "spreadsheet"
          ? "workbook"
          : "presentation",
    container: "zip",
  };
}

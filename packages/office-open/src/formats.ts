import { CompoundFileReader, OOXML_PACKAGE_FORMATS } from "@office-open/core";
import { unzipSync } from "fflate";

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

export function detectOfficeFormat(input: Uint8Array | string): OfficeFormatInfo {
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

  const text = new TextDecoder("latin1", { fatal: false }).decode(input.subarray(0, 64));
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
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(data);
  } catch {
    throw new Error("Unable to read the ZIP office container");
  }

  const mimeType = files["mimetype"];
  if (mimeType) {
    const format = ODF_MIME_TYPES[new TextDecoder().decode(mimeType).trim()];
    if (!format) throw new Error("Unsupported ODF mimetype");
    return {
      format,
      family: format === "ods" ? "workbook" : format === "odp" ? "presentation" : "document",
      container: "zip",
    };
  }

  const contentTypes = files["[Content_Types].xml"];
  if (!contentTypes) throw new Error("ZIP office container is missing [Content_Types].xml");
  const xml = new TextDecoder().decode(contentTypes);
  const match = Object.entries(OOXML_PACKAGE_FORMATS).find(([, info]) =>
    xml.includes(info.mainContentType),
  );
  if (!match) throw new Error("Unsupported OOXML package");
  const [format, info] = match as [OfficeFormat, (typeof match)[1]];
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

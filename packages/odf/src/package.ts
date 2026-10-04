import { parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const MANIFEST_COMPATIBILITY_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.3";

export type OdfFileContent = string | Uint8Array;

export interface OdfFiles {
  [path: string]: string;
}

export interface OdfPackageFiles {
  [path: string]: OdfFileContent;
}

export function manifestXml(mimeType: string, files: OdfPackageFiles): string {
  const paths = ["mimetype", ...directoryPaths(Object.keys(files)), ...Object.keys(files)];
  const entries = ["/", ...paths]
    .map(
      (path) =>
        `<manifest:file-entry manifest:full-path="${path}"${
          path === "/" ? ` manifest:version="1.3"` : ""
        } manifest:media-type="${path === "mimetype" ? mimeType : mediaType(path)}"/>`,
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="1.3">${entries}</manifest:manifest>`;
}

function mediaType(path: string): string {
  const imageTypes: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    tif: "image/tiff",
    svg: "image/svg+xml",
  };
  if (path.endsWith(".xml")) return "text/xml";
  return imageTypes[path.split(".").pop() ?? ""] ?? "application/binary";
}

function directoryPaths(paths: string[]): string[] {
  return [
    ...new Set(
      paths.flatMap((path) =>
        path
          .split("/")
          .slice(0, -1)
          .map((_, index, parts) => parts.slice(0, index + 1).join("/") + "/"),
      ),
    ),
  ];
}

export function generateOcf(mimeType: string, files: OdfPackageFiles): Uint8Array {
  return zipSync(
    {
      mimetype: [strToU8(mimeType), { level: 0 }],
      ...Object.fromEntries(
        Object.entries(files).map(([path, content]) => [
          path,
          typeof content === "string" ? strToU8(content) : [content, { level: 6 }],
        ]),
      ),
      "META-INF/manifest.xml": strToU8(manifestXml(mimeType, files)),
    },
    { level: 6 },
  );
}

export function readOcf(
  data: Uint8Array,
  expectedMimeType: string,
): { files: OdfFiles; binaries: Record<string, Uint8Array>; manifest: Element } {
  const entries = unzipSync(data);
  const files: OdfFiles = {};
  const binaries: Record<string, Uint8Array> = {};
  for (const [path, bytes] of Object.entries(entries)) {
    if (path === "mimetype") continue;
    if (path.endsWith(".xml")) {
      files[path] = strFromU8(bytes);
    } else {
      binaries[path] = bytes;
    }
  }
  const mimeType = strFromU8(entries.mimetype ?? new Uint8Array());
  if (mimeType !== expectedMimeType) {
    throw new Error(`Unexpected ODF MIME type: ${mimeType || "missing"}`);
  }
  const manifestXml = files["META-INF/manifest.xml"];
  if (!manifestXml) throw new Error("ODF package is missing META-INF/manifest.xml");
  const manifestDocument = parse(manifestXml, {
    ignoreDeclaration: true,
    ignoreDoctype: true,
    normalizeNamespaces: {
      [MANIFEST_NS]: "manifest",
      [MANIFEST_COMPATIBILITY_NS]: "manifest",
    },
  });
  const manifest = manifestDocument.elements?.[0] ?? manifestDocument;
  if (manifest.name !== "manifest:manifest") throw new Error("Invalid ODF manifest");
  validateManifestPaths(manifest, ["content.xml"]);
  return { files, binaries, manifest };
}

function validateManifestPaths(manifest: Element, actualPaths: string[]): void {
  const declared = new Set(
    (manifest.elements ?? [])
      .filter((element) => element.name === "manifest:file-entry")
      .map((element) => String(manifestAttribute(element, "full-path") ?? "")),
  );
  for (const path of actualPaths) {
    if (!declared.has(path)) throw new Error(`Manifest does not declare ${path}`);
  }
}

function manifestAttribute(element: Element, name: string): string | undefined {
  const value = element.attributes?.[`manifest:${name}`] ?? element.attributes?.[name];
  return value === undefined ? undefined : String(value);
}

export function readXml(files: OdfFiles, path: string): Element {
  const xml = files[path];
  if (!xml) throw new Error(`ODF package is missing ${path}`);
  const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
  return document.elements?.[0] ?? document;
}

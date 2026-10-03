import { parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.3";

export interface OdfFiles {
  [path: string]: string;
}

export function manifestXml(mimeType: string, files: OdfFiles): string {
  const paths = ["mimetype", ...Object.keys(files)];
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
  return path.endsWith(".xml") ? "text/xml" : "application/binary";
}

export function generateOcf(mimeType: string, files: OdfFiles): Uint8Array {
  return zipSync(
    {
      mimetype: [strToU8(mimeType), { level: 0 }],
      ...Object.fromEntries(Object.entries(files).map(([path, xml]) => [path, strToU8(xml)])),
      "META-INF/manifest.xml": strToU8(manifestXml(mimeType, files)),
    },
    { level: 6 },
  );
}

export function readOcf(
  data: Uint8Array,
  expectedMimeType: string,
): { files: OdfFiles; manifest: Element } {
  const entries = unzipSync(data);
  const files: OdfFiles = {};
  for (const [path, bytes] of Object.entries(entries)) {
    if (path !== "mimetype") files[path] = strFromU8(bytes);
  }
  const mimeType = strFromU8(entries.mimetype ?? new Uint8Array());
  if (mimeType !== expectedMimeType) {
    throw new Error(`Unexpected ODF MIME type: ${mimeType || "missing"}`);
  }
  const manifestXml = files["META-INF/manifest.xml"];
  if (!manifestXml) throw new Error("ODF package is missing META-INF/manifest.xml");
  const manifestDocument = parse(manifestXml, { ignoreDeclaration: true });
  const manifest = manifestDocument.elements?.[0] ?? manifestDocument;
  if (manifest.name !== "manifest:manifest") throw new Error("Invalid ODF manifest");
  validateManifestPaths(manifest, ["content.xml"]);
  return { files, manifest };
}

function validateManifestPaths(manifest: Element, actualPaths: string[]): void {
  const declared = new Set(
    (manifest.elements ?? [])
      .filter((element) => element.name === "manifest:file-entry")
      .map((element) => String(element.attributes?.["manifest:full-path"] ?? "")),
  );
  for (const path of actualPaths) {
    if (!declared.has(path)) throw new Error(`Manifest does not declare ${path}`);
  }
}

export function readXml(files: OdfFiles, path: string): Element {
  const xml = files[path];
  if (!xml) throw new Error(`ODF package is missing ${path}`);
  const document = parse(xml, { ignoreDeclaration: true });
  return document.elements?.[0] ?? document;
}

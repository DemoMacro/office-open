import { parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

import { OcfManifestError, OcfMimeTypeError, OdfXmlError } from "../errors";
import {
  manifestOptionsXml,
  parseManifestOptions,
  type OcfManifestFileEntryOptions,
  type OcfManifestOptions,
} from "./manifest";

export type OdfFileContent = string | Uint8Array;

export interface OdfFiles {
  [path: string]: string;
}

export interface OdfPackageFiles {
  [path: string]: OdfFileContent;
}

export function manifestXml(
  mimeType: string,
  files: OdfPackageFiles,
  mediaTypes: Record<string, string> = {},
  packageManifest?: OcfManifestOptions,
): string {
  const entries = canonicalManifestEntries(mimeType, files, mediaTypes, packageManifest);
  return `<?xml version="1.0" encoding="UTF-8"?>${manifestOptionsXml({
    version: packageManifest?.version ?? "1.3",
    entries,
    ...(packageManifest?.encryptedKeys?.length
      ? { encryptedKeys: packageManifest.encryptedKeys }
      : {}),
  })}`;
}

function canonicalManifestEntries(
  mimeType: string,
  files: OdfPackageFiles,
  mediaTypes: Record<string, string>,
  packageManifest?: OcfManifestOptions,
): OcfManifestFileEntryOptions[] {
  const filePaths = new Set(Object.keys(files));
  const declared = packageManifestEntries(packageManifest).filter((entry) =>
    isGeneratedManifestPath(entry.fullPath, filePaths),
  );
  const declaredRoot = declared.find((entry) => entry.fullPath === "/");
  if (declaredRoot?.mediaType !== undefined && declaredRoot.mediaType !== mimeType)
    throw new OcfManifestError(
      "Manifest root media type differs from package mimetype: " + declaredRoot.mediaType,
      { fullPath: "/" },
    );
  const byPath = new Map<string, OcfManifestFileEntryOptions>();
  for (const entry of declared) {
    if (byPath.has(entry.fullPath))
      throw new OcfManifestError(`Duplicate OCF manifest full-path: ${entry.fullPath}`, {
        fullPath: entry.fullPath,
      });
    byPath.set(entry.fullPath, entry);
  }
  const paths = ["mimetype", ...directoryPaths(Object.keys(files)), ...Object.keys(files)];
  for (const path of paths) {
    if (byPath.has(path)) continue;
    const sourceType =
      path === "mimetype" ? mimeType : (mediaTypes[path] ?? packageMediaType(declared, path));
    byPath.set(path, {
      fullPath: path,
      ...(path === "/" ? { version: "1.3" } : {}),
      mediaType: sourceType ?? mediaType(path),
    });
  }
  if (!byPath.has("/")) byPath.set("/", { fullPath: "/", version: "1.3", mediaType: mimeType });
  const declaredPaths = new Set(declared.map((entry) => entry.fullPath));
  const appended = ["/", ...paths]
    .filter((path) => !declaredPaths.has(path) && byPath.has(path))
    .map((path) => byPath.get(path)!);
  const entries = [...declared, ...appended];
  return entries;
}

function packageManifestEntries(
  packageManifest: OcfManifestOptions | undefined,
): OcfManifestFileEntryOptions[] {
  return packageManifest?.entries ?? [];
}

function packageMediaType(
  entries: readonly OcfManifestFileEntryOptions[],
  path: string,
): string | undefined {
  return entries.find((entry) => entry.fullPath === path)?.mediaType;
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

function isGeneratedManifestPath(path: string, filePaths: Set<string>): boolean {
  return path === "/" || path === "mimetype" || path.endsWith("/") || filePaths.has(path);
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

export function generateOcf(
  mimeType: string,
  files: OdfPackageFiles,
  mediaTypes: Record<string, string> = {},
  packageManifest?: OcfManifestOptions,
): Uint8Array {
  const declaredDirectories = new Set(
    (packageManifest?.entries ?? [])
      .map((entry) => entry.fullPath)
      .filter((path) => path.endsWith("/") && path !== "/"),
  );
  return zipSync(
    {
      mimetype: [strToU8(mimeType), { level: 0 }],
      ...Object.fromEntries(
        [...declaredDirectories].map((path) => [path, [new Uint8Array(), { level: 0 }]]),
      ),
      ...Object.fromEntries(
        Object.entries(files).map(([path, content]) => [
          path,
          typeof content === "string" ? strToU8(content) : [content, { level: 6 }],
        ]),
      ),
      "META-INF/manifest.xml": strToU8(manifestXml(mimeType, files, mediaTypes, packageManifest)),
    },
    { level: 6 },
  );
}

export function readOcf(
  data: Uint8Array,
  expectedMimeType: string,
): { files: OdfFiles; binaries: Record<string, Uint8Array>; manifest: OcfManifestOptions } {
  const entries = unzipSync(data);
  const files: OdfFiles = {};
  const binaries: Record<string, Uint8Array> = {};
  for (const [path, bytes] of Object.entries(entries)) {
    if (path === "mimetype") continue;
    if (path.endsWith("/")) continue;
    if (path.endsWith(".xml")) {
      files[path] = strFromU8(bytes);
    } else {
      binaries[path] = bytes;
    }
  }
  const manifestXml = files["META-INF/manifest.xml"];
  if (!manifestXml) {
    throw new OcfManifestError("ODF package is missing META-INF/manifest.xml");
  }
  let manifest: OcfManifestOptions;
  try {
    manifest = parseManifestOptions(manifestXml);
  } catch (cause) {
    if (cause instanceof OcfManifestError) throw cause;
    throw new OcfManifestError("Invalid ODF manifest", { cause });
  }
  const rootEntry = manifest.entries.find((entry) => entry.fullPath === "/");
  const mimeType = entries.mimetype ? strFromU8(entries.mimetype) : rootEntry?.mediaType;
  if (mimeType && mimeType !== expectedMimeType)
    throw new OcfMimeTypeError(expectedMimeType, mimeType);
  validateManifestPaths(
    manifest,
    Object.keys(entries).filter(
      (path) => !path.endsWith("/") && path !== "mimetype" && path !== "META-INF/manifest.xml",
    ),
    Object.keys(entries).filter((path) => path !== "mimetype" && path !== "META-INF/manifest.xml"),
  );
  return { files, binaries, manifest };
}

export function hasOcfManifestOverlay(manifest: OcfManifestOptions): boolean {
  return (
    (manifest.encryptedKeys?.length ?? 0) > 0 ||
    manifest.entries.some(
      (entry) =>
        entry.preferredViewMode !== undefined ||
        entry.size !== undefined ||
        entry.encryptionData !== undefined ||
        (entry.version !== undefined && entry.fullPath !== "/") ||
        (entry.fullPath.endsWith("/") && entry.fullPath !== "/"),
    )
  );
}

function validateManifestPaths(
  manifest: OcfManifestOptions,
  actualPaths: string[],
  retainedPaths: string[] = actualPaths,
): void {
  const declared = new Set(manifest.entries.map((entry) => entry.fullPath));
  for (const path of actualPaths) {
    if (!declared.has(path)) {
      manifest.entries.push({ fullPath: path, mediaType: mediaType(path) });
    }
  }
  const actualSet = new Set(actualPaths);
  manifest.entries = manifest.entries.filter((entry) => {
    if (entry.fullPath === "/" || entry.fullPath === "mimetype") return true;
    const directoryPath = entry.fullPath.endsWith("/") ? entry.fullPath : `${entry.fullPath}/`;
    return entry.fullPath.endsWith("/")
      ? actualSet.has(entry.fullPath) ||
          retainedPaths.some((path) => path.startsWith(directoryPath))
      : actualSet.has(entry.fullPath);
  });
}

export function readXml(files: OdfFiles, path: string): Element {
  const xml = files[path];
  if (!xml) throw new OdfXmlError(path);
  try {
    const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
    return document.elements?.[0] ?? document;
  } catch (cause) {
    throw new OdfXmlError(path, { cause });
  }
}

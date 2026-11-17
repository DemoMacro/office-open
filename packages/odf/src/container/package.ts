import type { DataType } from "@office-open/core";
import { ParsedArchive, strFromU8, strToU8, toUint8ArrayAsync, zipSync } from "@office-open/core";
import { parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { OcfError, OcfManifestError, OcfMimeTypeError, OdfXmlError } from "../errors";
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

const ODF_ZIP_READ_LIMITS = {
  maxEntries: 20_000,
  maxEntryUncompressedSize: 256 * 1024 * 1024,
  maxTotalUncompressedSize: 1024 * 1024 * 1024,
} as const;

const ODF_CANONICAL_PREFIXES: Readonly<Record<string, string>> = {
  "http://openoffice.org/2009/office": "officeooo",
  "http://openoffice.org/2009/table": "tableooo",
  "http://www.gnumeric.org/odf-extension/1.0": "gnm",
  "http://www.w3.org/1999/02/22-rdf-syntax-ns#": "rdf",
  "http://www.w3.org/1999/xlink": "xlink",
  "http://www.w3.org/2002/xforms": "xforms",
  "http://www.w3.org/XML/1998/namespace": "xml",
  "http://purl.org/dc/elements/1.1/": "dc",
  "urn:oasis:names:tc:opendocument:xmlns:animation:1.0": "anim",
  "urn:oasis:names:tc:opendocument:xmlns:chart:1.0": "chart",
  "urn:oasis:names:tc:opendocument:xmlns:config:1.0": "config",
  "urn:oasis:names:tc:opendocument:xmlns:database:1.0": "db",
  "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0": "number",
  "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0": "draw",
  "urn:oasis:names:tc:opendocument:xmlns:form:1.0": "form",
  "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0": "manifest",
  "urn:oasis:names:tc:opendocument:xmlns:manifest:1.3": "manifest",
  "urn:oasis:names:tc:opendocument:xmlns:meta:1.0": "meta",
  "urn:oasis:names:tc:opendocument:xmlns:office:1.0": "office",
  "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0": "presentation",
  "urn:oasis:names:tc:opendocument:xmlns:script:1.0": "script",
  "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0": "smil",
  "urn:oasis:names:tc:opendocument:xmlns:style:1.0": "style",
  "urn:oasis:names:tc:opendocument:xmlns:table:1.0": "table",
  "urn:oasis:names:tc:opendocument:xmlns:text:1.0": "text",
  "urn:oasis:names:tc:opendocument:xmlns:version:1.0": "version",
  "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0": "fo",
  "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0": "svg",
  "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0": "calcext",
  "urn:org:documentfoundation:names:experimental:chart:xmlns:loext:1.0": "chartooo",
  "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0": "loext",
};

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
  assertManifestPaths(packageManifest?.entries ?? []);
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
  assertPackagePaths(files);
  assertManifestPaths(packageManifest?.entries ?? []);
  const declaredDirectories = new Set(
    (packageManifest?.entries ?? [])
      .map((entry) => entry.fullPath)
      .filter((path) => path.endsWith("/") && path !== "/"),
  );
  assertNoPackagePathConflicts(
    ["mimetype", "META-INF/manifest.xml", ...Object.keys(files)],
    [...declaredDirectories],
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

export function readOcfSync(
  data: Uint8Array,
  expectedMimeType: string | readonly string[],
): {
  files: OdfFiles;
  binaries: Record<string, Uint8Array>;
  manifest: OcfManifestOptions;
  mimeType?: string;
} {
  const expectedMimeTypes = Array.isArray(expectedMimeType) ? expectedMimeType : [expectedMimeType];
  let entries: Record<string, Uint8Array>;
  try {
    const archive = new ParsedArchive(data, undefined, ODF_ZIP_READ_LIMITS);
    entries = Object.fromEntries(archive.keys().map((path) => [path, archive.getRaw(path)!]));
  } catch (cause) {
    throw new OcfError("Invalid ODF ZIP package", { cause });
  }
  assertNoPackagePathConflicts(
    Object.keys(entries).filter((path) => !path.endsWith("/")),
    Object.keys(entries).filter((path) => path.endsWith("/")),
  );
  const files: OdfFiles = {};
  const binaries: Record<string, Uint8Array> = {};
  for (const [path, bytes] of Object.entries(entries)) {
    if (path === "mimetype" || path === "META-INF/manifest.xml" || path.endsWith("/")) continue;
    if (path.endsWith(".xml")) {
      files[path] = strFromU8(bytes);
    } else {
      Object.defineProperty(binaries, path, {
        configurable: true,
        enumerable: true,
        value: bytes,
        writable: true,
      });
    }
  }
  const manifestData = entries["META-INF/manifest.xml"];
  const manifestXml = manifestData ? strFromU8(manifestData) : undefined;
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
  manifest = {
    ...manifest,
    entries: [
      ...manifest.entries,
      ...Object.keys(entries)
        .filter(
          (path) =>
            !path.endsWith("/") &&
            path !== "mimetype" &&
            path !== "META-INF/manifest.xml" &&
            !manifest.entries.some((entry) => entry.fullPath === path),
        )
        .map((path) => ({
          fullPath: path,
          mediaType: packageMediaType(manifest.entries, path) ?? mediaType(path),
        })),
    ],
  };
  const rootEntry = manifest.entries.find((entry) => entry.fullPath === "/");
  const mimeType = entries.mimetype ? strFromU8(entries.mimetype) : rootEntry?.mediaType;
  if (
    rootEntry?.mediaType !== undefined &&
    mimeType !== undefined &&
    rootEntry.mediaType !== mimeType
  ) {
    throw new OcfMimeTypeError(rootEntry.mediaType, mimeType);
  }
  if (mimeType && !expectedMimeTypes.includes(mimeType))
    throw new OcfMimeTypeError(expectedMimeTypes[0]!, mimeType);
  validateManifestPaths(
    manifest,
    Object.keys(entries).filter(
      (path) => !path.endsWith("/") && path !== "mimetype" && path !== "META-INF/manifest.xml",
    ),
    Object.keys(entries).filter((path) => path !== "mimetype" && path !== "META-INF/manifest.xml"),
  );
  return { files, binaries, manifest, mimeType };
}

export async function readOcf(
  data: DataType,
  expectedMimeType: string | readonly string[],
): Promise<{
  files: OdfFiles;
  binaries: Record<string, Uint8Array>;
  manifest: OcfManifestOptions;
  mimeType?: string;
}> {
  return readOcfSync(await toUint8ArrayAsync(data), expectedMimeType);
}

export function hasOcfManifestOptions(manifest: OcfManifestOptions): boolean {
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
  const declared = new Set<string>();
  for (const entry of manifest.entries) {
    const path = entry.fullPath;
    if (!isSafePackagePath(path, true)) {
      throw new OcfManifestError(`Invalid ODF manifest path: ${path}`, { fullPath: path });
    }
    if (declared.has(path)) {
      throw new OcfManifestError(`Duplicate ODF manifest path: ${path}`, { fullPath: path });
    }
    declared.add(path);
  }
  const actualSet = new Set(actualPaths);
  for (const path of actualPaths) {
    if (!declared.has(path)) {
      throw new OcfManifestError(`ODF manifest does not declare package member: ${path}`, {
        fullPath: path,
      });
    }
  }
  for (const entry of manifest.entries) {
    if (
      entry.fullPath === "/" ||
      entry.fullPath === "mimetype" ||
      entry.fullPath === "META-INF/manifest.xml"
    )
      continue;
    const directoryPath = entry.fullPath.endsWith("/") ? entry.fullPath : `${entry.fullPath}/`;
    if (entry.fullPath.endsWith("/")) {
      if (
        !actualSet.has(entry.fullPath) &&
        !retainedPaths.some((path) => path.startsWith(directoryPath))
      ) {
        continue;
      }
    } else if (isOptionalMissingManifestPath(entry.fullPath)) {
      continue;
    } else if (!actualSet.has(entry.fullPath)) {
      throw new OcfManifestError(
        `ODF manifest declares missing package member: ${entry.fullPath}`,
        { fullPath: entry.fullPath },
      );
    }
  }
}

function assertPackagePaths(files: OdfPackageFiles): void {
  for (const path of Object.keys(files)) {
    if (
      !isSafePackagePath(path, false) ||
      path.endsWith("/") ||
      path.toLowerCase() === "mimetype" ||
      path.toLowerCase() === "meta-inf/manifest.xml"
    ) {
      throw new OcfManifestError(`Invalid ODF package member path: ${path}`, {
        fullPath: path,
      });
    }
  }
}

function assertManifestPaths(entries: readonly OcfManifestFileEntryOptions[]): void {
  for (const { fullPath } of entries) {
    if (!isSafePackagePath(fullPath, true)) {
      throw new OcfManifestError(`Invalid ODF manifest path: ${fullPath}`, { fullPath });
    }
  }
}

function isSafePackagePath(path: string, allowRoot: boolean): boolean {
  if (allowRoot && path === "/") return true;
  if (!path || path.startsWith("/") || path.includes("\\")) return false;
  const segments = path.endsWith("/") ? path.slice(0, -1).split("/") : path.split("/");
  return (
    segments.length > 0 &&
    !/^[a-z]:/i.test(segments[0]!) &&
    segments.every((segment) => segment !== "" && segment !== "." && segment !== "..")
  );
}

function assertNoPackagePathConflicts(
  filePaths: readonly string[],
  directoryPaths: readonly string[],
): void {
  const files = new Set(filePaths);
  const directories = new Set(directoryPaths.map((path) => path.slice(0, -1)));
  for (const path of [...files, ...directories]) {
    const segments = path.split("/");
    while (segments.length > 1) {
      segments.pop();
      const parent = segments.join("/");
      if (files.has(parent)) {
        throw new OcfManifestError(`ODF package file/directory path conflict: ${parent}`, {
          fullPath: parent,
        });
      }
    }
  }
  for (const path of files) {
    if (directories.has(path)) {
      throw new OcfManifestError(`ODF package file/directory path conflict: ${path}`, {
        fullPath: path,
      });
    }
  }
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

export function readXml(files: OdfFiles, path: string): Element {
  const xml = files[path];
  if (!xml) throw new OdfXmlError(path);
  try {
    const document = parse(xml, {
      ignoreDeclaration: true,
      ignoreDoctype: true,
      normalizeNamespaces: ODF_CANONICAL_PREFIXES,
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
      throw new Error("XML part must contain exactly one root element");
    }
    return roots[0]!;
  } catch (cause) {
    throw new OdfXmlError(path, { cause });
  }
}

import {
  convertOutput,
  toUint8Array,
  toUint8ArrayAsync,
  streamUint8Array,
  type ChartSpaceOptions,
  type DataType,
  type OfficeFormsOptions,
  type OutputByType,
  type OutputType,
} from "@office-open/core";
import {
  attributeString,
  CHART_MIME,
  chartSubdocumentFiles,
  childNamed,
  childrenNamed,
  emuToLength,
  descendantElements,
  generateOcf,
  hasOcfManifestOverlay,
  isExternalOdfUrl,
  lengthToEmu,
  metaXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  officeFormsXml,
  parseOfficeForms,
  parseGraphicStyles,
  parseMeta,
  readOcfSync,
  readXml,
  xmlElement,
  type OcfManifestOptions,
  type OdfPackageFiles,
} from "@office-open/odf";
import type { OdfPackageMemberOptions } from "@office-open/odf";
import type { PresentationOptions, SlideOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import type { OdpChart, OdpImage } from "./drawing";
import type { OdpEmbeddedObjectPool } from "./embedded-frames";
import { OdpParseError } from "./error";
import type { OdpPresentationOptions } from "./options";
import { parseColumnWidths, parseSlide, slideXml } from "./slide";
import { parseStyleOverlays, parseTextStyles, styleOverlaysXml } from "./styles";

export const MIME = "application/vnd.oasis.opendocument.presentation";
export const TEMPLATE_MIME = "application/vnd.oasis.opendocument.presentation-template";

export const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:xforms="http://www.w3.org/2002/xforms"',
  'xmlns:officeooo="http://openoffice.org/2009/office"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(" ");

function isOdpOverlay(options: PresentationOptions): options is OdpPresentationOptions {
  return (
    "packageManifest" in options ||
    "packageMembers" in options ||
    "styleOverlays" in options ||
    "embeddedObjects" in options
  );
}

function packageOverlays(options: PresentationOptions): Partial<OdpPresentationOptions> {
  return isOdpOverlay(options) ? options : {};
}

function generateOdpPackage(options: PresentationOptions): Uint8Array {
  const { packageManifest, packageMembers, styleOverlays, embeddedObjects } =
    packageOverlays(options);
  validateEmbeddedObjects(options, embeddedObjects);
  const styles: string[] = styleOverlaysXml(styleOverlays);
  const images: OdpImage[] = [];
  const mediaFiles: OdpImage[] = [];
  const charts: OdpChart[] = [];
  const embeddedObjectPool = new Map(
    (embeddedObjects ?? []).map((object) => [object.path ?? object.sourceUrl!, object] as const),
  );
  const size = normalizeSize(options.size);
  const pageLayout = xmlElement("style:page-layout", { "style:name": "PM1" }, [
    xmlElement("style:page-layout-properties", {
      "fo:page-width": emuToLength(size.width),
      "fo:page-height": emuToLength(size.height),
    }),
  ]);
  const pages = (options.slides ?? []).map((slide, index) =>
    slideXml(slide, index + 1, styles, images, mediaFiles, embeddedObjectPool, charts),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(pages.join(""), styles),
    "styles.xml": stylesXml(pageLayout),
    "meta.xml": metaXml(options),
  };
  const mediaTypes: Record<string, string> = Object.fromEntries(
    charts.map((chart) => [`${chart.path}/`, CHART_MIME]),
  );
  for (const object of embeddedObjects ?? []) {
    if (object.path) {
      validateObjectPath(object.path);
      if (object.mediaType !== undefined) mediaTypes[`${object.path}/`] = object.mediaType;
    }
    for (const member of object.members) {
      addPackageFile(files, member);
      if (member.mediaType !== undefined) mediaTypes[member.path] = member.mediaType;
    }
  }
  for (const member of packageMembers ?? []) addPackageFile(files, member);
  for (const image of images) files[image.path] = image.data;
  for (const media of mediaFiles) files[media.path] = media.data;
  for (const chart of charts)
    Object.assign(files, chartSubdocumentFiles(`${chart.path}/`, chart.chart));
  return generateOcf(MIME, files, mediaTypes, packageManifest);
}

function validateEmbeddedObjects(
  options: PresentationOptions,
  embeddedObjects: OdpPresentationOptions["embeddedObjects"],
): void {
  const keys = new Set<string>();
  for (const object of embeddedObjects ?? []) {
    if ((object.path === undefined) === (object.sourceUrl === undefined))
      throw new Error("ODP embedded object must have exactly one package path or source URL");
    const key = object.path ?? object.sourceUrl!;
    if (keys.has(key)) throw new Error(`Duplicate ODP embedded object: ${key}`);
    keys.add(key);
  }

  const shapeCounts = new Map<number, number>();
  const countShapes = (children: SlideOptions["children"]): void => {
    for (const child of children ?? []) {
      if ("shape" in child && child.shape.id !== undefined)
        shapeCounts.set(child.shape.id, (shapeCounts.get(child.shape.id) ?? 0) + 1);
      if ("group" in child) countShapes(child.group.children);
    }
  };
  for (const slide of options.slides ?? []) countShapes(slide.children);
  for (const object of embeddedObjects ?? []) {
    for (const frame of object.frames ?? []) {
      if (shapeCounts.get(frame.shapeId) !== 1) {
        throw new Error(
          `ODP embedded object frame does not resolve to one shape: ${frame.shapeId}`,
        );
      }
    }
  }
}

function validateObjectPath(path: string): void {
  if (!path || path.startsWith("/") || path.endsWith("/") || path.split("/").includes(".."))
    throw new Error(`Invalid ODP embedded object path: ${path}`);
}

function addPackageFile(files: OdfPackageFiles, member: OdfPackageMemberOptions): void {
  if (
    !member.path ||
    member.path.startsWith("/") ||
    member.path.endsWith("/") ||
    member.path.split("/").includes("..")
  )
    throw new Error(`Invalid ODP package member path: ${member.path}`);
  if (member.path === "META-INF/manifest.xml")
    throw new Error("ODP package members cannot replace META-INF/manifest.xml");
  if (files[member.path] !== undefined)
    throw new Error(`ODP package member conflicts with modeled content: ${member.path}`);
  files[member.path] = typeof member.data === "string" ? member.data : toUint8Array(member.data);
}

export function generatePresentationSync<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  output?: T,
): OutputByType[T] {
  return convertOutput(generateOdpPackage(options), (output ?? "nodebuffer") as T, MIME);
}

export async function generatePresentation<T extends OutputType = "nodebuffer">(
  options: PresentationOptions,
  output?: T,
): Promise<OutputByType[T]> {
  return generatePresentationSync(options, output);
}

export function generatePresentationStream(
  options: PresentationOptions,
): ReadableStream<Uint8Array> {
  return streamUint8Array(generateOdpPackage(options));
}

export async function parsePresentation(data: DataType): Promise<OdpPresentationOptions> {
  return parsePresentationSync(await toUint8ArrayAsync(data));
}

export function parsePresentationSync(data: DataType): OdpPresentationOptions {
  try {
    return parseOdpPackage(toUint8Array(data));
  } catch (cause) {
    if (cause instanceof OdpParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdpParseError(cause.message, cause.part, cause.path, cause.name, cause.reason, {
        cause,
      });
    }
    throw new OdpParseError(
      cause instanceof Error ? cause.message : "Unable to parse ODP package",
      "mimetype",
      "/mimetype",
      "mimetype",
      "invalid ODP package",
      { cause },
    );
  }
}

function parseOdpPackage(data: Uint8Array): OdpPresentationOptions {
  const { files, binaries, manifest } = readOcfSync(data, [MIME, TEMPLATE_MIME]);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:presentation");
  const stylesDocument = files["styles.xml"] ? readXml(files, "styles.xml") : undefined;
  const pageLayout = childNamed(
    childNamed(childNamed(stylesDocument, "office:automatic-styles"), "style:page-layout"),
    "style:page-layout-properties",
  );
  const width = lengthToEmu(attributeString(pageLayout, "fo:page-width"));
  const height = lengthToEmu(attributeString(pageLayout, "fo:page-height"));
  const graphicStyles = parseGraphicStyles(childNamed(content, "office:automatic-styles"));
  const chartPool = parseEmbeddedCharts(manifest, files, binaries);
  const embeddedObjects = embeddedObjectPool(manifest, files, binaries, chartPool);
  const manifestMediaTypes = new Map(
    manifest.entries.map((entry) => [entry.fullPath, entry.mediaType] as const),
  );
  const usedBinaryPaths = new Set<string>();
  validateEmbeddedObjectBodies(body, binaries, chartPool, embeddedObjects);
  const nextEmbeddedShapeId = { value: 1 };
  const result = {
    ...parseMeta(files),
    ...(width && height ? { size: { width, height } } : {}),
    slides: childrenNamed(body, "draw:page").map((page) =>
      parseSlide(
        page,
        parseTextStyles(childNamed(content, "office:automatic-styles")),
        parseColumnWidths(childNamed(content, "office:automatic-styles")),
        binaries,
        graphicStyles,
        chartPool,
        usedBinaryPaths,
        embeddedObjects,
        manifestMediaTypes,
        nextEmbeddedShapeId,
      ),
    ),
  };
  const styleOverlays = parseStyleOverlays(childNamed(content, "office:automatic-styles"));
  const packageMembers = sourcePackageMembers(
    manifest,
    files,
    binaries,
    chartPool,
    usedBinaryPaths,
    embeddedObjects,
  );
  return {
    ...result,
    ...(hasOcfManifestOverlay(manifest) ? { packageManifest: manifest } : {}),
    ...(packageMembers.length > 0 ? { packageMembers } : {}),
    ...(embeddedObjects.size > 0 ? { embeddedObjects: [...embeddedObjects.values()] } : {}),
    ...(styleOverlays.length > 0 ? { styleOverlays } : {}),
  };
}

function validateEmbeddedObjectBodies(
  body: Element | undefined,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
  embeddedObjects: OdpEmbeddedObjectPool,
): void {
  const references = (body ? descendantElements(body) : []).flatMap((element) =>
    element.name === "draw:object-ole"
      ? [attributeString(element, "xlink:href")?.replace(/^\.\//, "").replace(/^\//, "")]
      : element.name === "draw:object"
        ? [attributeString(element, "xlink:href")?.replace(/^\.\//, "").replace(/^\//, "")]
        : [],
  );
  for (const href of references) {
    if (
      href &&
      (binaries[href] ||
        chartPool.has(href) ||
        embeddedObjects.has(href) ||
        isExternalEmbeddedHref(href))
    )
      continue;
    throw new OdpParseError(
      `content.xml: ${href ?? ""}: embedded object subdocument is missing`,
      "content.xml",
      "/draw:page/draw:frame/draw:object/@xlink:href",
      "draw:object",
      "embedded-object-subdocument-missing",
    );
  }
}

function embeddedObjectPool(
  manifest: OcfManifestOptions,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
): OdpEmbeddedObjectPool {
  const normalizePath = (path: string): string =>
    decodeURIComponent(
      path
        .replace(/^\.\//, "")
        .replace(/^\//, "")
        .replace(/\/$/, "")
        .replace(/%(?![0-9A-Fa-f]{2})/g, "%25"),
    );
  const mediaTypes = new Map(
    manifest.entries
      .filter((entry) => entry.fullPath.endsWith("/") && entry.fullPath !== "/")
      .map((entry) => [entry.fullPath.replace(/\/$/, ""), entry.mediaType] as const),
  );
  const pool: OdpEmbeddedObjectPool = new Map();
  for (const [rawPath, mediaType] of mediaTypes) {
    const path = normalizePath(rawPath);
    if (chartPool.has(path) || files[`${rawPath}/content.xml`] === undefined) continue;
    pool.set(path, {
      path,
      ...(mediaType === undefined ? {} : { mediaType }),
      members: objectMembers(path, files, binaries),
    });
  }
  return pool;
}

function objectMembers(
  path: string,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
): OdfPackageMemberOptions[] {
  return [
    ...Object.entries(files)
      .filter(([memberPath]) => memberPath.startsWith(`${path}/`))
      .map(([memberPath, data]) => ({ path: memberPath, data })),
    ...Object.entries(binaries)
      .filter(([memberPath]) => memberPath.startsWith(`${path}/`))
      .map(([memberPath, data]) => ({ path: memberPath, data })),
  ].sort((left, right) => left.path.localeCompare(right.path));
}

function isExternalEmbeddedHref(href: string): boolean {
  return isExternalOdfUrl(href) || href.startsWith("///");
}

function sourcePackageMembers(
  manifest: OcfManifestOptions,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
  usedBinaryPaths: Set<string>,
  embeddedObjects: OdpEmbeddedObjectPool,
): OdfPackageMemberOptions[] {
  const objectMemberPaths = new Set(
    [...embeddedObjects.values()].flatMap((object) => object.members.map((member) => member.path)),
  );
  const mediaTypes = new Map(
    manifest.entries.map((entry) => [entry.fullPath, entry.mediaType] as const),
  );
  const modeledPaths = new Set([
    "content.xml",
    "styles.xml",
    "meta.xml",
    "META-INF/manifest.xml",
    ...[...chartPool.keys()].map((path) => `${path}/content.xml`),
    ...usedBinaryPaths,
  ]);
  return [
    ...Object.entries(files).map(([path, data]) => ({ path, data })),
    ...Object.entries(binaries).map(([path, data]) => ({ path, data })),
  ]
    .filter(({ path }) => !modeledPaths.has(path) && !objectMemberPaths.has(path))
    .sort((left, right) => left.path.localeCompare(right.path))
    .map(({ path, data }) => ({ path, mediaType: mediaTypes.get(path), data }));
}

export function contentXml(pages: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
}

export function stylesXml(pageLayout: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/><office:automatic-styles>${pageLayout}</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
}

export function normalizeSize(size: PresentationOptions["size"]): {
  width: number;
  height: number;
} {
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  if (typeof size === "object" && size !== null) {
    return {
      width: typeof size.width === "number" ? size.width : (lengthToEmu(size.width) ?? 0),
      height: typeof size.height === "number" ? size.height : (lengthToEmu(size.height) ?? 0),
    };
  }
  return { width: 12192000, height: 6858000 };
}

export function formsXml(forms: OfficeFormsOptions | undefined): string {
  return officeFormsXml(forms);
}

export function parseForms(page: Element): OfficeFormsOptions | undefined {
  const element = childNamed(page, "office:forms");
  return element === undefined ? undefined : parseOfficeForms(element);
}

export function unknownSlideChild(element: Element, parentPath: string): OdpParseError {
  const name = element.name ?? "(unknown)";
  const path = `${parentPath}/${name}`;
  const reason = "element has no canonical PresentationOptions mapping";
  return new OdpParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}

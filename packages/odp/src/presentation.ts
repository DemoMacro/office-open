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
  type SyncDataType,
} from "@office-open/core";
import {
  attributeString,
  bindGraphicStyleResources,
  CHART_MIME,
  chartSubdocumentFiles,
  childNamed,
  childrenNamed,
  createGraphicStyleResources,
  emuToLength,
  descendantElements,
  generateOcf,
  graphicStyleFiles,
  graphicStyleResourcesXml,
  hexColorValue,
  hasOcfManifestOptions,
  isExternalOdfUrl,
  lengthToEmu,
  metaXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  officeFormsXml,
  parseOfficeForms,
  parseGraphicStyles,
  parseGraphicStyleResources,
  parseMeta,
  readOcfSync,
  readXml,
  xmlElement,
  type OcfManifestOptions,
  type GraphicStyle,
  type OdfPackageFiles,
  type StyleResourceOptions,
} from "@office-open/odf";
import type { PackageMemberOptions } from "@office-open/odf";
import type { MasterDefinition, PresentationOptions, SlideOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import type { OdpChart, OdpImage } from "./drawing";
import { normalizeEmbeddedObjectPath, type OdpEmbeddedObjectPool } from "./embedded-frames";
import { OdpGenerateError, OdpParseError } from "./error";
import type {
  OdpMasterDefinition,
  OdpMasterHeaderFooterOptions,
  OdpPageLayoutOptions,
  OdpPresentationOptions,
  OdpSlideOptions,
} from "./options";
import {
  parseColumnWidths,
  parseSlide,
  reconcileTransitionStyleResources,
  slideChildXml,
  slideXml,
} from "./slide";
import { parseStyleResources, parseTextStyles, styleResourcesXml } from "./styles";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";
import { materializeOdpTransitionSound, type OdpTransitionOptions } from "./transition";

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

function isOdpPackageOptions(options: PresentationOptions): options is OdpPresentationOptions {
  return (
    "packageManifest" in options ||
    "packageMembers" in options ||
    "styleResources" in options ||
    "embeddedObjects" in options ||
    "packageMediaType" in options
  );
}

function packageOptions(options: PresentationOptions): Partial<OdpPresentationOptions> {
  return isOdpPackageOptions(options) ? options : {};
}

function backgroundOptions(
  slides: readonly OdpSlideOptions[],
  resources: readonly StyleResourceOptions[],
): { slides: OdpSlideOptions[]; resources: StyleResourceOptions[] } {
  const nextResources = [...resources];
  const usedNames = new Set(resources.map((resource) => resource.name));
  const nextSlides = slides.map((slide, index) => {
    if (slide.background === undefined) return slide;
    const background = slide.background;
    if (
      background.reference !== undefined ||
      background.effects !== undefined ||
      background.effectsList !== undefined ||
      background.shadeToTitle !== undefined ||
      background.blackWhiteMode !== undefined
    )
      throw new OdpGenerateError(
        "ODF drawing-page styles cannot represent these background properties",
        "content.xml",
        "/office:document-content/office:automatic-styles/style:style/style:drawing-page-properties",
        "background",
        "unsupported-background-properties",
      );
    const fill = background.fill;
    let fillMode: string;
    let fillColor: string | undefined;
    if (fill === undefined || (typeof fill === "object" && fill.type === "none")) {
      fillMode = "none";
    } else if (
      typeof fill === "object" &&
      fill.type === "solid" &&
      hexColorValue(fill.color) !== undefined
    ) {
      const color = hexColorValue(fill.color)!.slice(-6);
      if (!/^[0-9a-f]{6}$/iu.test(color))
        throw new OdpGenerateError(
          "ODF slide background requires an RGB color",
          "content.xml",
          "/office:document-content/office:automatic-styles/style:style/style:drawing-page-properties",
          "background fill color",
          "invalid-rgb-color",
        );
      fillMode = "solid";
      fillColor = `#${color}`;
    } else {
      throw new OdpGenerateError(
        "ODF slide background supports only solid colors and no-fill",
        "content.xml",
        "/office:document-content/office:automatic-styles/style:style/style:drawing-page-properties",
        "background fill",
        "unsupported-background-fill",
      );
    }
    let name = slide.drawingPageStyleName;
    if (name === undefined) {
      let suffix = index + 1;
      name = `ODPBackground${suffix}`;
      while (usedNames.has(name)) name = `ODPBackground${++suffix}`;
    }
    usedNames.add(name);
    const resourceIndex = nextResources.findIndex(
      (resource) => resource.name === name && resource.family === "drawing-page",
    );
    const resource = resourceIndex < 0 ? undefined : nextResources[resourceIndex];
    const properties = [...(resource?.properties ?? [])];
    const propertyIndex = properties.findIndex(
      (property) => property.name === "style:drawing-page-properties",
    );
    const existingProperty = properties[propertyIndex];
    const drawingPageProperties = {
      name: "style:drawing-page-properties" as const,
      attributes: {
        ...existingProperty?.attributes,
        "draw:fill": fillMode,
        "draw:fill-color": fillColor,
      },
      ...(existingProperty && "children" in existingProperty && existingProperty.children
        ? { children: existingProperty.children }
        : {}),
    };
    if (propertyIndex < 0) properties.push(drawingPageProperties);
    else properties[propertyIndex] = drawingPageProperties;
    const nextResource: StyleResourceOptions = {
      name,
      family: "drawing-page",
      properties,
      ...(resource?.maps === undefined ? {} : { maps: resource.maps }),
    };
    if (resourceIndex < 0) nextResources.push(nextResource);
    else nextResources[resourceIndex] = nextResource;
    return { ...slide, drawingPageStyleName: name };
  });
  return { slides: nextSlides, resources: nextResources };
}

function generateOdpPackage(options: OdpPresentationOptions): Uint8Array {
  const { packageManifest, packageMembers, styleResources, embeddedObjects, packageMediaType } =
    packageOptions(options);
  validateEmbeddedObjects(options, embeddedObjects);
  const images: OdpImage[] = [];
  const mediaFiles: OdpImage[] = [];
  const background = backgroundOptions(options.slides ?? [], styleResources ?? []);
  const slides = background.slides;
  const usedStyleNames = new Set(background.resources.map((resource) => resource.name));
  const occupiedPaths = new Set([
    ...(packageMembers ?? []).map((member) => member.path),
    ...(embeddedObjects ?? []).flatMap((object) => object.members.map((member) => member.path)),
  ]);
  const transitionSounds = new Map<OdpSlideOptions, OdpTransitionOptions>();
  for (const slide of slides) {
    if (slide.transition === undefined || typeof slide.transition === "string") continue;
    transitionSounds.set(
      slide,
      materializeOdpTransitionSound(slide.transition, mediaFiles, occupiedPaths),
    );
  }
  const presentationOptions = { ...options, slides };
  const styles: string[] = styleResourcesXml(
    reconcileTransitionStyleResources(
      { ...presentationOptions, styleResources: background.resources },
      transitionSounds,
    ),
  );
  const graphicResources = createGraphicStyleResources();
  bindGraphicStyleResources(styles, graphicResources);
  const masterNames = [
    ...(options.masters ?? []).flatMap((master) => (master.name ? [master.name] : [])),
    ...slides.flatMap((slide) => (slide.master ? [slide.master] : [])),
  ].filter((name, index, names) => names.indexOf(name) === index);
  if (masterNames.length === 0) masterNames.push("Default");
  const charts: OdpChart[] = [];
  const embeddedObjectPool = new Map(
    (embeddedObjects ?? []).map((object) => [object.path ?? object.sourceUrl!, object] as const),
  );
  const size = normalizeSize(options.size);
  const pageLayouts = options.pageLayouts?.length
    ? options.pageLayouts
    : [{ name: "PM1", size: options.size ?? { width: size.width, height: size.height } }];
  const pageLayout = pageLayouts.map((layout) => pageLayoutXml(layout, size)).join("");
  const masterPages = masterPagesXml(
    masterNames,
    options.masters ?? [],
    pageLayouts[0]?.name ?? "PM1",
    styles,
    images,
    mediaFiles,
    embeddedObjectPool,
    charts,
  );
  const pages = slides.map((slide, index) =>
    slideXml(
      slide,
      index + 1,
      styles,
      images,
      mediaFiles,
      embeddedObjectPool,
      charts,
      usedStyleNames,
      transitionSounds,
      occupiedPaths,
    ),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(pages.join(""), styles),
    "styles.xml": stylesXml(pageLayout, masterPages, graphicStyleResourcesXml(graphicResources)),
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
  for (const file of graphicStyleFiles(graphicResources)) files[file.path] = file.data;
  for (const media of mediaFiles) files[media.path] = media.data;
  for (const chart of charts)
    Object.assign(files, chartSubdocumentFiles(`${chart.path}/`, chart.chart));
  return generateOcf(packageMediaType ?? MIME, files, mediaTypes, packageManifest);
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

function addPackageFile(files: OdfPackageFiles, member: PackageMemberOptions): void {
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

export function generatePresentationSync<T extends OutputType = "uint8array">(
  options: OdpPresentationOptions,
  output?: T,
): OutputByType[T] {
  const mediaType = options.packageMediaType ?? MIME;
  return convertOutput(generateOdpPackage(options), (output ?? "uint8array") as T, mediaType);
}

export async function generatePresentation<T extends OutputType = "uint8array">(
  options: OdpPresentationOptions,
  output?: T,
): Promise<OutputByType[T]> {
  return generatePresentationSync(options, output);
}

export function generatePresentationStream(
  options: OdpPresentationOptions,
): ReadableStream<Uint8Array> {
  return streamUint8Array(generateOdpPackage(options));
}

export async function parsePresentation(data: DataType): Promise<OdpPresentationOptions> {
  return parsePresentationSync(await toUint8ArrayAsync(data));
}

export function parsePresentationSync(data: SyncDataType): OdpPresentationOptions {
  try {
    return parseOdpPackage(toUint8Array(data));
  } catch (cause) {
    if (cause instanceof OdpParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdpParseError(
        cause.message,
        cause.part,
        cause.path,
        cause.elementName,
        cause.reason,
        { cause },
      );
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
  if (!body) {
    throw new OdpParseError(
      "content.xml: /office:document-content/office:body/office:presentation: required presentation body is missing",
      "content.xml",
      "/office:document-content/office:body/office:presentation",
      "office:presentation",
      "missing-required-element",
    );
  }
  const stylesDocument = files["styles.xml"] ? readXml(files, "styles.xml") : undefined;
  const pageLayouts = parsePageLayouts(stylesDocument);
  const masterStyles = childNamed(stylesDocument, "office:master-styles");
  const firstMaster = childNamed(masterStyles, "style:master-page");
  const firstLayoutName = attributeString(firstMaster, "style:page-layout-name");
  const firstPageLayout =
    pageLayouts.find((layout) => layout.name === firstLayoutName) ?? pageLayouts[0];
  const size = firstPageLayout?.size;
  const chartPool = parseEmbeddedCharts(manifest, files, binaries, files["styles.xml"]);
  const embeddedObjects = embeddedObjectPool(manifest, files, binaries, chartPool);
  const manifestMediaTypes = new Map(
    manifest.entries.map((entry) => [entry.fullPath, entry.mediaType] as const),
  );
  const usedBinaryPaths = new Set<string>();
  validateEmbeddedObjectBodies(body, binaries, chartPool, embeddedObjects);
  const nextEmbeddedShapeId = { value: 1 };
  const contentAutomaticStyles = childNamed(content, "office:automatic-styles");
  const stylesAutomaticStyles = childNamed(stylesDocument, "office:automatic-styles");
  const commonStyles = childNamed(stylesDocument, "office:styles");
  const graphicResources = parseGraphicStyleResources(
    [contentAutomaticStyles, stylesAutomaticStyles, commonStyles],
    binaries,
  );
  const styleResources = [
    ...parseStyleResources(contentAutomaticStyles),
    ...parseStyleResources(stylesAutomaticStyles),
    ...parseStyleResources(commonStyles),
  ]
    .filter(
      (resource, index, resources) =>
        resources.findIndex(
          (candidate) => candidate.name === resource.name && candidate.family === resource.family,
        ) === index,
    )
    .filter((resource) => resource.family !== "graphic");
  const commonGraphicStyles = parseGraphicStyles(commonStyles, "graphic", graphicResources);
  const stylesGraphicStyles = parseGraphicStyles(
    stylesAutomaticStyles,
    "graphic",
    graphicResources,
    commonGraphicStyles,
  );
  const graphicStyles = new Map([
    ...commonGraphicStyles,
    ...stylesGraphicStyles,
    ...parseGraphicStyles(
      contentAutomaticStyles,
      "graphic",
      graphicResources,
      new Map([...commonGraphicStyles, ...stylesGraphicStyles]),
    ),
  ]);
  const textStyles = new Map([
    ...parseTextStyles(contentAutomaticStyles),
    ...parseTextStyles(stylesAutomaticStyles),
    ...parseTextStyles(commonStyles),
  ]);
  const columnWidths = new Map([
    ...parseColumnWidths(contentAutomaticStyles),
    ...parseColumnWidths(stylesAutomaticStyles),
    ...parseColumnWidths(commonStyles),
  ]);
  const masterStyleResources = [
    ...parseStyleResources(contentAutomaticStyles),
    ...parseStyleResources(stylesAutomaticStyles),
    ...parseStyleResources(commonStyles),
  ];
  const masters = childrenNamed(masterStyles, "style:master-page").flatMap((master) => {
    const name = attributeString(master, "style:name");
    if (name === undefined) return [];
    const children = parseMasterChildren(
      master,
      textStyles,
      columnWidths,
      binaries,
      files,
      graphicStyles,
      chartPool,
      usedBinaryPaths,
      embeddedObjects,
      manifestMediaTypes,
      nextEmbeddedShapeId,
      masterStyleResources,
    );
    if (childNamed(master, "loext:theme")) {
      throw new OdpParseError(
        "styles.xml: ODF master theme extension has no ODF 1.3 canonical mapping",
        "styles.xml",
        "/office:document-styles/office:master-styles/style:master-page/loext:theme",
        "loext:theme",
        "unsupported-odf-theme",
      );
    }
    return [
      {
        name,
        ...(attributeString(master, "style:page-layout-name")
          ? { pageLayoutName: attributeString(master, "style:page-layout-name") }
          : {}),
        ...(children.length > 0 ? { children } : {}),
        ...parseMasterHeaderFooter(master, textStyles),
      },
    ];
  });
  const result = {
    ...parseMeta(files),
    ...(size ? { size } : {}),
    slides: childrenNamed(body, "draw:page").map((page) =>
      parseSlide(
        page,
        textStyles,
        columnWidths,
        binaries,
        files,
        graphicStyles,
        chartPool,
        usedBinaryPaths,
        embeddedObjects,
        manifestMediaTypes,
        nextEmbeddedShapeId,
        styleResources,
      ),
    ),
  };
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
    ...(masters.length > 0 ? { masters } : {}),
    ...(pageLayouts.length > 0 ? { pageLayouts } : {}),
    ...(hasOcfManifestOptions(manifest) ? { packageManifest: manifest } : {}),
    ...(packageMembers.length > 0 ? { packageMembers } : {}),
    ...(embeddedObjects.size > 0 ? { embeddedObjects: [...embeddedObjects.values()] } : {}),
    ...(styleResources.length > 0 ? { styleResources } : {}),
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
      ? [normalizeEmbeddedObjectPath(attributeString(element, "xlink:href") ?? "")]
      : element.name === "draw:object"
        ? [normalizeEmbeddedObjectPath(attributeString(element, "xlink:href") ?? "")]
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
  const mediaTypes = new Map(
    manifest.entries
      .filter((entry) => entry.fullPath.endsWith("/") && entry.fullPath !== "/")
      .map((entry) => [entry.fullPath.replace(/\/$/, ""), entry.mediaType] as const),
  );
  const pool: OdpEmbeddedObjectPool = new Map();
  for (const [rawPath, mediaType] of mediaTypes) {
    const path = normalizeEmbeddedObjectPath(rawPath, "META-INF/manifest.xml");
    if (chartPool.has(path) || files[`${rawPath}/content.xml`] === undefined) continue;
    pool.set(path, {
      path,
      ...(mediaType === undefined ? {} : { mediaType }),
      members: objectMembers(rawPath, files, binaries),
    });
  }
  return pool;
}

function objectMembers(
  path: string,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
): PackageMemberOptions[] {
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
): PackageMemberOptions[] {
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

export function stylesXml(
  pageLayout: string,
  masterPages: readonly string[] = [
    '<style:master-page style:name="Default" style:page-layout-name="PM1"/>',
  ],
  graphicResources: readonly string[] = [],
): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles>${graphicResources.join("")}</office:styles><office:automatic-styles>${pageLayout}</office:automatic-styles><office:master-styles>${masterPages.join("")}</office:master-styles></office:document-styles>`;
}

function pageLayoutXml(
  layout: OdpPageLayoutOptions,
  fallback: { width: number; height: number },
): string {
  const size = layout.size === undefined ? fallback : normalizeSize(layout.size);
  return xmlElement("style:page-layout", { "style:name": layout.name }, [
    xmlElement("style:page-layout-properties", {
      "fo:page-width": emuToLength(size.width),
      "fo:page-height": emuToLength(size.height),
    }),
  ]);
}

function masterPagesXml(
  names: readonly string[],
  masters: readonly OdpMasterDefinition[],
  defaultPageLayoutName: string,
  styles: string[],
  images: OdpImage[],
  mediaFiles: OdpImage[],
  embeddedObjects: OdpEmbeddedObjectPool,
  charts: OdpChart[],
): string[] {
  return names.map((name) => {
    const master = masters.find((candidate) => candidate.name === name);
    if (master?.theme) {
      throw new OdpGenerateError(
        `ODF 1.3 has no master theme element for "${name}"`,
        "styles.xml",
        `/office:document-styles/office:master-styles/style:master-page[@style:name="${name}"]`,
        "theme",
        "unsupported-odf-theme",
      );
    }
    const children = (master?.children ?? []).map((child) => {
      if ("placeholder" in child || "placeholderShape" in child) {
        throw new OdpGenerateError(
          `ODP master placeholder cannot be generated from a PPTX-only placeholder reference: ${name}`,
          "styles.xml",
          `/office:document-styles/office:master-styles/style:master-page[@style:name="${name}"]`,
          "placeholder",
          "unsupported-master-placeholder",
        );
      }
      return slideChildXml(child, styles, images, mediaFiles, embeddedObjects, charts);
    });
    return xmlElement(
      "style:master-page",
      {
        "style:name": name,
        "style:page-layout-name": master?.pageLayoutName ?? defaultPageLayoutName,
      },
      [...masterHeaderFooterXml(master, styles), ...children],
    );
  });
}

function masterHeaderFooterXml(
  master: OdpMasterDefinition | undefined,
  styles: string[],
): string[] {
  if (!master) return [];
  const regions: [string, OdpMasterHeaderFooterOptions | undefined][] = [
    ["header", master.header],
    ["header-left", master.headerLeft],
    ["header-first", master.headerFirst],
    ["footer", master.footer],
    ["footer-left", master.footerLeft],
    ["footer-first", master.footerFirst],
  ];
  return regions.flatMap(([name, value]) => {
    if (value === undefined) return [];
    return [
      xmlElement(
        `style:${name}`,
        { "style:display": value.display },
        value.paragraphs === undefined ? [] : textBodyXml({ paragraphs: value.paragraphs }, styles),
      ),
    ];
  });
}

function parseMasterChildren(
  master: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
  binaries: Record<string, Uint8Array>,
  files: OdfPackageFiles,
  graphicStyles: Map<string, GraphicStyle>,
  chartPool: Map<string, ChartSpaceOptions>,
  usedBinaryPaths: Set<string>,
  embeddedObjects: OdpEmbeddedObjectPool,
  manifestMediaTypes: Map<string, string | undefined>,
  nextEmbeddedShapeId: { value: number },
  styleResources: StyleResourceOptions[],
): NonNullable<MasterDefinition["children"]> {
  const page: Element = {
    ...master,
    name: "draw:page",
    elements: (master.elements ?? []).filter(
      (child) =>
        child.name !== "loext:theme" &&
        child.name !== "style:header" &&
        child.name !== "style:header-left" &&
        child.name !== "style:header-first" &&
        child.name !== "style:footer" &&
        child.name !== "style:footer-left" &&
        child.name !== "style:footer-first",
    ),
  };
  try {
    return (
      parseSlide(
        page,
        textStyles,
        columnWidths,
        binaries,
        files,
        graphicStyles,
        chartPool,
        usedBinaryPaths,
        embeddedObjects,
        manifestMediaTypes,
        nextEmbeddedShapeId,
        styleResources,
      ).children ?? []
    );
  } catch (error) {
    if (!(error instanceof OdpParseError) || error.part !== "content.xml") throw error;
    const path = error.path
      .replace(
        "/office:document-content/office:body/office:presentation",
        "/office:document-styles/office:master-styles",
      )
      .replace("/draw:page", "/style:master-page");
    throw new OdpParseError(
      error.message.replace("content.xml:", "styles.xml:"),
      "styles.xml",
      path,
      error.elementName,
      error.reason,
      { cause: error },
    );
  }
}

function parseMasterHeaderFooter(
  master: Element,
  textStyles: Map<string, TextProperties>,
): Pick<
  OdpMasterDefinition,
  "header" | "headerLeft" | "headerFirst" | "footer" | "footerLeft" | "footerFirst"
> {
  const result: Pick<
    OdpMasterDefinition,
    "header" | "headerLeft" | "headerFirst" | "footer" | "footerLeft" | "footerFirst"
  > = {};
  const regions = [
    ["style:header", "header"],
    ["style:header-left", "headerLeft"],
    ["style:header-first", "headerFirst"],
    ["style:footer", "footer"],
    ["style:footer-left", "footerLeft"],
    ["style:footer-first", "footerFirst"],
  ] as const;
  for (const [elementName, optionName] of regions) {
    const element = childNamed(master, elementName);
    if (!element) continue;
    const path = `/office:document-styles/office:master-styles/style:master-page/${elementName}`;
    const unsupported = Object.keys(element.attributes ?? {}).find(
      (attribute) => attribute !== "style:display",
    );
    if (unsupported) throw unsupportedMasterHeaderFooter(element, path, unsupported);
    const children = (element.elements ?? []).filter((child) => child.type === "element");
    const unsupportedChild = children.find((child) => child.name !== "text:p");
    if (unsupportedChild)
      throw unsupportedMasterHeaderFooter(element, path, unsupportedChild.name!);
    result[optionName] = {
      ...(attributeBoolean(element, "style:display") === undefined
        ? {}
        : { display: attributeBoolean(element, "style:display") }),
      ...(children.length > 0
        ? { paragraphs: children.map((child) => parseParagraph(child, textStyles)) }
        : {}),
    };
  }
  return result;
}

function attributeBoolean(element: Element, name: string): boolean | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  if (value === "true" || value === "1") return true;
  if (value === "false" || value === "0") return false;
  throw unsupportedMasterHeaderFooter(
    element,
    `/office:document-styles/office:master-styles/style:master-page/${element.name}`,
    name,
  );
}

function unsupportedMasterHeaderFooter(
  element: Element,
  path: string,
  name: string,
): OdpParseError {
  return new OdpParseError(
    `styles.xml: ${path}: unsupported master header/footer member ${name}`,
    "styles.xml",
    `${path}/${name}`,
    element.name ?? "style:header",
    "unsupported-master-header-footer",
  );
}

function parsePageLayouts(stylesDocument: Element | undefined): OdpPageLayoutOptions[] {
  return childrenNamed(
    childNamed(stylesDocument, "office:automatic-styles"),
    "style:page-layout",
  ).flatMap((layout) => {
    const name = attributeString(layout, "style:name");
    if (name === undefined) return [];
    const properties = childNamed(layout, "style:page-layout-properties");
    const width = lengthToEmu(attributeString(properties, "fo:page-width"));
    const height = lengthToEmu(attributeString(properties, "fo:page-height"));
    if (width === undefined && height === undefined) return [{ name }];
    if (width === undefined || height === undefined) {
      throw new OdpParseError(
        `styles.xml: page layout "${name}" has only one page dimension`,
        "styles.xml",
        `/office:document-styles/office:automatic-styles/style:page-layout[@style:name="${name}"]/style:page-layout-properties`,
        "style:page-layout-properties",
        "incomplete-page-layout",
      );
    }
    return [{ name, size: { width, height } }];
  });
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

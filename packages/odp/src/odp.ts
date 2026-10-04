import type {
  ChartSpaceOptions,
  ParagraphDescriptorOptions,
  TextBodyOptions,
  TextRunOptions,
} from "@office-open/core";
import { toUint8Array, type EndpointConnectionOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  generateOcf,
  lengthToEmu,
  metaXml,
  parseMeta,
  readOcf,
  readXml,
  textOf,
  xmlElement,
  type OdfPackageFiles,
} from "@office-open/ocf";
import {
  CHART_MIME,
  chartBodyXml,
  graphicFill,
  graphicOutline,
  OdfSchemaError,
  parseEmbeddedCharts,
  parseGraphicStyles,
  PRESET_GEOMETRY_DOCX,
  presetGeometryOdf,
  pushShapeStyle,
} from "@office-open/odf-schema";
import type { GraphicStyle } from "@office-open/odf-schema";
import type {
  ChartOptions,
  ConnectorOptions,
  LineShapeOptions,
  PictureOptions,
  PresentationOptions,
  ShapeOptions,
  SlideChild,
  SlideOptions,
  TableOptions,
} from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";

const MIME = "application/vnd.oasis.opendocument.presentation";
const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
].join(" ");

interface TextProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: "single";
  size?: number;
}

/** Binary image collected during generation — emitted as a Pictures/ entry. */
interface OdpImage {
  path: string;
  data: Uint8Array;
}

/** Embedded chart subdocument collected during generation. */
interface OdpChart {
  path: string;
  chart: ChartOptions;
}

export function generateOdp(options: PresentationOptions): Uint8Array {
  const styles: string[] = [];
  const images: OdpImage[] = [];
  const charts: OdpChart[] = [];
  const size = normalizeSize(options.size);
  const pageLayout = xmlElement("style:page-layout", { "style:name": "PM1" }, [
    xmlElement("style:page-layout-properties", {
      "fo:page-width": emuToLength(size.width),
      "fo:page-height": emuToLength(size.height),
    }),
  ]);
  const pages = (options.slides ?? []).map((slide, index) =>
    slideXml(slide, index + 1, styles, images, charts),
  );
  const files: OdfPackageFiles = {
    "content.xml": contentXml(pages.join(""), styles),
    "styles.xml": stylesXml(pageLayout),
    "meta.xml": metaXml(options),
  };
  for (const image of images) files[image.path] = image.data;
  for (const chart of charts) files[`${chart.path}/content.xml`] = chartBodyXml(chart.chart);
  return generateOcf(
    MIME,
    files,
    Object.fromEntries(charts.map((chart) => [`${chart.path}/`, CHART_MIME])),
  );
}

export function parseOdp(data: Uint8Array): PresentationOptions {
  try {
    return parseOdpPresentation(data);
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

function parseOdpPresentation(data: Uint8Array): PresentationOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
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
  const chartPool = parseEmbeddedCharts(manifest, files);
  return {
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
      ),
    ),
  };
}

function contentXml(pages: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
}

function stylesXml(pageLayout: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/><office:automatic-styles>${pageLayout}</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
}

function normalizeSize(size: PresentationOptions["size"]): { width: number; height: number } {
  if (size === "4:3") return { width: 9144000, height: 6858000 };
  if (typeof size === "object" && size !== null) {
    return {
      width: typeof size.width === "number" ? size.width : (lengthToEmu(size.width) ?? 0),
      height: typeof size.height === "number" ? size.height : (lengthToEmu(size.height) ?? 0),
    };
  }
  return { width: 12192000, height: 6858000 };
}

function slideXml(
  slide: SlideOptions,
  index: number,
  styles: string[],
  images: OdpImage[],
  charts: OdpChart[],
): string {
  const frames = (slide.children ?? []).map((child) =>
    slideChildXml(child, styles, images, charts),
  );
  const notes = typeof slide.notes === "string" ? slide.notes : slide.notes?.text;
  const notesXml = notes
    ? xmlElement("presentation:notes", undefined, [
        xmlElement("draw:frame", undefined, [
          xmlElement("draw:text-box", undefined, [
            xmlElement("text:p", undefined, [escapeText(notes)]),
          ]),
        ]),
      ])
    : undefined;
  return xmlElement(
    "draw:page",
    { "draw:name": `Slide${index}`, "draw:master-page-name": "Default" },
    [...frames, ...(notesXml ? [notesXml] : [])],
  );
}

/** Recursive SlideChild → ODF dispatcher shared by slides and draw:g groups. */
function slideChildXml(
  child: SlideChild,
  styles: string[],
  images: OdpImage[],
  charts: OdpChart[],
): string {
  if ("shape" in child) {
    if (child.shape.properties?.geometry !== undefined) return customShapeXml(child.shape, styles);
    return shapeXml(child.shape, styles);
  }
  if ("table" in child) return slideTableXml(child.table, styles);
  if ("connector" in child) return connectorXml(child.connector);
  if ("line" in child) return lineXml(child.line);
  if ("picture" in child) return pictureFrameXml(child.picture, images);
  if ("chart" in child) return chartFrameXml(child.chart, charts);
  if ("group" in child)
    return xmlElement(
      "draw:g",
      { "draw:name": child.group.name },
      child.group.children.map((nested) => slideChildXml(nested, styles, images, charts)),
    );
  return "";
}

/** Shape with preset geometry renders as draw:custom-shape + enhanced-geometry. */
function customShapeXml(shape: ShapeOptions, styles: string[]): string {
  const geometry = shape.properties?.geometry;
  const preset =
    geometry === undefined ? undefined : typeof geometry === "string" ? geometry : geometry.preset;
  return xmlElement(
    "draw:custom-shape",
    {
      "draw:style-name": pushShapeStyle(
        shape.properties?.fill ?? undefined,
        shape.properties?.outline ?? undefined,
        styles,
      ),
      "svg:x": toOdfLength(shape.x),
      "svg:y": toOdfLength(shape.y),
      "svg:width": toOdfLength(shape.width),
      "svg:height": toOdfLength(shape.height),
      "draw:name": shape.name,
    },
    [
      xmlElement(
        "draw:enhanced-geometry",
        { "draw:type": preset ? presetGeometryOdf(preset) : undefined },
        [],
      ),
      ...textBodyXml(shape.textBody, styles),
    ],
  );
}

function shapeXml(shape: ShapeOptions, styles: string[]): string {
  return xmlElement(
    "draw:frame",
    {
      "draw:name": shape.name,
      "svg:x": emuToLength(shape.x),
      "svg:y": emuToLength(shape.y),
      "svg:width": emuToLength(shape.width),
      "svg:height": emuToLength(shape.height),
    },
    [xmlElement("draw:text-box", undefined, textBodyXml(shape.textBody, styles))],
  );
}

function textBodyXml(body: TextBodyOptions | undefined, styles: string[]): string[] {
  const paragraphs = body?.paragraphs ?? (body?.text !== undefined ? [body.text] : [""]);
  return paragraphs.map((paragraph) => {
    const value = typeof paragraph === "string" ? { text: paragraph } : paragraph;
    const children =
      value.text !== undefined && value.children === undefined
        ? [escapeText(value.text)]
        : (value.children ?? []).map((run) => {
            if (typeof run === "string")
              return xmlElement("text:span", undefined, [escapeText(run)]);
            if (!("text" in run)) return "";
            const properties = textProperties(run);
            const styleName = addTextStyle(properties, styles);
            return xmlElement("text:span", { "text:style-name": styleName }, [
              escapeText(run.text ?? ""),
            ]);
          });
    return xmlElement("text:p", undefined, children);
  });
}

/** Serializes a connector as a straight draw:connector with endpoint gluing. */
function connectorXml(connector: ConnectorOptions): string {
  return xmlElement("draw:connector", {
    "draw:type": "line",
    "draw:name": connector.name,
    "svg:x1": toOdfLength(connector.x1),
    "svg:y1": toOdfLength(connector.y1),
    "svg:x2": toOdfLength(connector.x2),
    "svg:y2": toOdfLength(connector.y2),
    "draw:start-shape": connector.startConnection
      ? String(connector.startConnection.id)
      : undefined,
    "draw:start-glue-point": connector.startConnection?.index,
    "draw:end-shape": connector.endConnection ? String(connector.endConnection.id) : undefined,
    "draw:end-glue-point": connector.endConnection?.index,
  });
}

/** Serializes a straight line as draw:line. */
function lineXml(line: LineShapeOptions): string {
  return xmlElement("draw:line", {
    "draw:name": line.name,
    "svg:x1": toOdfLength(line.x1),
    "svg:y1": toOdfLength(line.y1),
    "svg:x2": toOdfLength(line.x2),
    "svg:y2": toOdfLength(line.y2),
  });
}

/** pptx EMU numbers pass through; universal measures resolve to EMU first. */
function toOdfLength(value: number | string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return emuToLength(typeof value === "number" ? value : (lengthToEmu(value) ?? 0));
}

/** Slide picture renders as a positioned draw:frame + draw:image. */
function pictureFrameXml(picture: PictureOptions, images: OdpImage[]): string {
  if (picture.data === undefined) return "";
  const data = toUint8Array(picture.data);
  const path = `Pictures/image${images.length + 1}.${picture.type}`;
  images.push({ path, data });
  return xmlElement(
    "draw:frame",
    {
      "draw:name": picture.name,
      "svg:x": toOdfLength(picture.x),
      "svg:y": toOdfLength(picture.y),
      "svg:width": toOdfLength(picture.width),
      "svg:height": toOdfLength(picture.height),
    },
    [xmlElement("draw:image", { "xlink:href": path })],
  );
}

/** Slide chart renders as a positioned draw:frame + embedded chart object. */
function chartFrameXml(chart: ChartOptions, charts: OdpChart[]): string {
  const path = `Object ${charts.length + 1}`;
  charts.push({ path, chart });
  return xmlElement(
    "draw:frame",
    {
      "draw:name": chart.name,
      "svg:x": toOdfLength(chart.x),
      "svg:y": toOdfLength(chart.y),
      "svg:width": toOdfLength(chart.width),
      "svg:height": toOdfLength(chart.height),
    },
    [
      xmlElement("draw:object", {
        "xlink:href": `./${path}`,
        "xlink:type": "simple",
        "xlink:show": "embed",
        "xlink:actuate": "onLoad",
      }),
    ],
  );
}

function textProperties(run: TextRunOptions): TextProperties {
  return {
    bold: run.bold,
    italic: run.italic,
    underline: run.underline === "single" ? "single" : undefined,
    size: run.size,
  };
}

function addTextStyle(properties: TextProperties, styles: string[]): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
      }),
    ]),
  );
  return name;
}

function parseSlide(
  page: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
  binaries: Record<string, Uint8Array>,
  graphicStyles: Map<string, GraphicStyle>,
  chartPool: Map<string, ChartSpaceOptions>,
): SlideOptions {
  const notes = childNamed(
    childNamed(childNamed(page, "presentation:notes"), "draw:frame"),
    "draw:text-box",
  );
  const notesText = notes ? textOf(childNamed(notes, "text:p")) : undefined;
  return {
    children:
      page.elements?.flatMap((child): SlideChild[] => {
        if (child.name === "draw:frame") {
          const chart = parseChartFrame(child, chartPool);
          if (chart) return [chart];
          const picture = parsePictureFrame(child, binaries);
          if (picture) return [picture];
          if (childNamed(child, "draw:text-box")) return [{ shape: parseShape(child, textStyles) }];
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        }
        if (child.name === "draw:custom-shape")
          return [{ shape: parseCustomSlideShape(child, textStyles, graphicStyles) }];
        if (child.name === "draw:line") return [parseLine(child)];
        if (child.name === "draw:g")
          return parseGroup(child, textStyles, columnWidths, binaries, chartPool);
        if (child.name === "draw:connector") return [parseConnector(child)];
        if (child.name === "table:table") return [parseSlideTable(child, textStyles, columnWidths)];
        if (child.name !== "presentation:notes")
          throw unknownSlideChild(
            child,
            "/office:document-content/office:body/office:presentation/draw:page",
          );
        return [];
      }) ?? [],
    ...(notesText ? { notes: notesText } : {}),
  };
}

/** Maps table-column style names to twip widths for slide tables. */
function parseColumnWidths(container: Element | undefined): Map<string, number> {
  const result = new Map<string, number>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-column") continue;
    const width = attributeString(
      childNamed(style, "style:table-column-properties"),
      "style:column-width",
    );
    if (width?.endsWith("cm")) {
      result.set(attributeString(style, "style:name") ?? "", lengthToEmu(width) ?? 0);
    }
  }
  return result;
}

/** Serializes a pptx table as an ODF table:table (cells render text paragraphs). */
function slideTableXml(table: TableOptions, styles: string[]): string {
  const columns = (table.columnWidths ?? []).map((width) =>
    xmlElement("table:table-column", { "table:style-name": addColumnStyle(width, styles) }),
  );
  const rows = table.rows.map((row) =>
    xmlElement(
      "table:table-row",
      undefined,
      row.cells.map((cell) => {
        const content = textBodyXml(
          cell.children ? { paragraphs: cell.children } : { text: cell.text ?? "" },
          styles,
        );
        const span = cell.columnSpan;
        const xml = xmlElement(
          "table:table-cell",
          { "office:value-type": "string", "table:number-columns-spanned": span },
          content,
        );
        const covered = Array.from({ length: Math.max(0, (span ?? 1) - 1) }, () =>
          xmlElement("table:covered-table-cell"),
        );
        return xml + covered.join("");
      }),
    ),
  );
  return xmlElement("table:table", { "table:name": `Table${styles.length + 1}` }, [
    ...columns,
    ...rows,
  ]);
}

/** Column width style — pptx widths are EMU. */
function addColumnStyle(width: number | string, styles: string[]): string {
  const name = `TC${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table-column" }, [
      xmlElement("style:table-column-properties", {
        "style:column-width": typeof width === "number" ? emuToLength(width) : width,
      }),
    ]),
  );
  return name;
}

/** Parses a slide table:table back to a pptx table child. */
function parseSlideTable(
  element: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
): { table: TableOptions } {
  return {
    table: {
      columnWidths: childrenNamed(element, "table:table-column").map(
        (column) => columnWidths.get(attributeString(column, "table:style-name") ?? "") ?? 0,
      ),
      rows: childrenNamed(element, "table:table-row").map((row) => ({
        cells: childrenNamed(row, "table:table-cell").map((cell) => ({
          columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
          children: childrenNamed(cell, "text:p").map((paragraph) =>
            parseParagraph(paragraph, textStyles),
          ),
        })),
      })),
    },
  };
}

function parseShape(frame: Element, textStyles: Map<string, TextProperties>): ShapeOptions {
  const textBox = childNamed(frame, "draw:text-box");
  return {
    name: attributeString(frame, "draw:name"),
    x: lengthToEmu(attributeString(frame, "svg:x")),
    y: lengthToEmu(attributeString(frame, "svg:y")),
    width: lengthToEmu(attributeString(frame, "svg:width")),
    height: lengthToEmu(attributeString(frame, "svg:height")),
    textBody: {
      paragraphs: childrenNamed(textBox, "text:p").map((paragraph) =>
        parseParagraph(paragraph, textStyles),
      ),
    },
  };
}

/** draw:custom-shape maps to the shared pptx shape with preset geometry. */
function parseCustomSlideShape(
  element: Element,
  textStyles: Map<string, TextProperties>,
  graphicStyles: Map<string, GraphicStyle>,
): ShapeOptions {
  const graphic = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const enhanced = childNamed(element, "draw:enhanced-geometry");
  const presetType = attributeString(enhanced, "draw:type");
  const preset = presetType ? PRESET_GEOMETRY_DOCX[presetType] : undefined;
  const paragraphs = childrenNamed(element, "text:p");
  const fill = graphicFill(graphic);
  const outline = graphicOutline(graphic);
  return {
    name: attributeString(element, "draw:name"),
    x: lengthToEmu(attributeString(element, "svg:x")),
    y: lengthToEmu(attributeString(element, "svg:y")),
    width: lengthToEmu(attributeString(element, "svg:width")),
    height: lengthToEmu(attributeString(element, "svg:height")),
    ...(paragraphs.length > 0
      ? {
          textBody: {
            paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)),
          },
        }
      : {}),
    ...(preset || fill || outline
      ? {
          properties: {
            ...(preset ? { geometry: preset } : {}),
            ...(fill ? { fill } : {}),
            ...(outline ? { outline } : {}),
          },
        }
      : {}),
  };
}

/** Parses a draw:connector back to a pptx connector child. */
function parseConnector(element: Element): { connector: ConnectorOptions } {
  return {
    connector: {
      x1: lengthToEmu(attributeString(element, "svg:x1")),
      y1: lengthToEmu(attributeString(element, "svg:y1")),
      x2: lengthToEmu(attributeString(element, "svg:x2")),
      y2: lengthToEmu(attributeString(element, "svg:y2")),
      startConnection: endpointConnection(element, "start"),
      endConnection: endpointConnection(element, "end"),
    },
  };
}

/** Parses a draw:line back to a pptx line child. */
function parseLine(element: Element): { line: LineShapeOptions } {
  return {
    line: {
      x1: lengthToEmu(attributeString(element, "svg:x1")),
      y1: lengthToEmu(attributeString(element, "svg:y1")),
      x2: lengthToEmu(attributeString(element, "svg:x2")),
      y2: lengthToEmu(attributeString(element, "svg:y2")),
    },
  };
}

/** Parses a draw:g group back to a pptx group child (children only). */
function parseGroup(
  element: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
): SlideChild[] {
  const children = (element.elements ?? []).flatMap((child): SlideChild[] => {
    if (child.name === "draw:frame") {
      const chart = parseChartFrame(child, chartPool);
      if (chart) return [chart];
      const picture = parsePictureFrame(child, binaries);
      if (picture) return [picture];
      if (childNamed(child, "draw:text-box")) return [{ shape: parseShape(child, textStyles) }];
      throw unknownSlideChild(child, "/draw:g");
    }
    if (child.name === "draw:line") return [parseLine(child)];
    if (child.name === "draw:connector") return [parseConnector(child)];
    if (child.name === "draw:g")
      return parseGroup(child, textStyles, columnWidths, binaries, chartPool);
    if (child.name === "table:table") return [parseSlideTable(child, textStyles, columnWidths)];
    throw unknownSlideChild(child, "/draw:g");
  });
  return [
    {
      group: {
        name: attributeString(element, "draw:name"),
        children,
      },
    },
  ];
}

/** Rejects unrecognized slide content instead of emitting generic XML. */
function unknownSlideChild(element: Element, parentPath: string): OdpParseError {
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

/** Parses a draw:frame + embedded chart subdocument back to a slide chart. */
function parseChartFrame(
  frame: Element,
  chartPool: Map<string, ChartSpaceOptions>,
): { chart: ChartOptions } | undefined {
  const object = childNamed(frame, "draw:object");
  if (!object) return undefined;
  const href = attributeString(object, "xlink:href")?.replace(/^\.\//, "").replace(/\/$/, "");
  const name = attributeString(frame, "draw:name") ?? href;
  const path = `/draw:frame[@draw:name="${name ?? ""}"]/draw:object`;
  if (!href || !name) {
    const reason = "chart frame has no object reference";
    throw new OdpParseError(
      `content.xml: ${path}: draw:object: ${reason}`,
      "content.xml",
      path,
      "draw:object",
      reason,
    );
  }
  const chart = chartPool.get(href);
  if (!chart) {
    const reason = "referenced chart subdocument is missing";
    throw new OdpParseError(
      `content.xml: ${path}: ${href}: ${reason}`,
      "content.xml",
      path,
      href,
      reason,
    );
  }
  return {
    chart: {
      ...chart,
      name,
      x: lengthToEmu(attributeString(frame, "svg:x")),
      y: lengthToEmu(attributeString(frame, "svg:y")),
      width: lengthToEmu(attributeString(frame, "svg:width")),
      height: lengthToEmu(attributeString(frame, "svg:height")),
    },
  };
}

/** Reads a draw:{start,end}-shape/-glue-point pair as a glued endpoint. */
function endpointConnection(
  element: Element,
  end: "start" | "end",
): EndpointConnectionOptions | undefined {
  const id = Number(attributeString(element, `draw:${end}-shape`));
  const index = attributeNumber(element, `draw:${end}-glue-point`);
  return Number.isInteger(id) && id > 0 && index !== undefined ? { id, index } : undefined;
}

/** Parses a draw:frame + draw:image back to a slide picture child. */
function parsePictureFrame(
  frame: Element,
  binaries: Record<string, Uint8Array>,
): { picture: PictureOptions } | undefined {
  const image = childNamed(frame, "draw:image");
  if (!image) return undefined;
  const path = attributeString(image, "xlink:href")?.replace(/^\//, "");
  const data = path ? binaries[path] : undefined;
  if (!data || !path) return undefined;
  return {
    picture: {
      type: (path.split(".").pop() ?? "png") as PictureOptions["type"],
      data,
      x: lengthToEmu(attributeString(frame, "svg:x")),
      y: lengthToEmu(attributeString(frame, "svg:y")),
      width: lengthToEmu(attributeString(frame, "svg:width")),
      height: lengthToEmu(attributeString(frame, "svg:height")),
    },
  };
}

function parseTextStyles(container: Element | undefined): Map<string, TextProperties> {
  const result = new Map<string, TextProperties>();
  for (const style of childrenNamed(container, "style:style")) {
    const properties = childNamed(style, "style:text-properties");
    const size = attributeString(properties, "fo:font-size");
    result.set(attributeString(style, "style:name") ?? "", {
      bold: attributeString(properties, "fo:font-weight") === "bold",
      italic: attributeString(properties, "fo:font-style") === "italic",
      underline:
        attributeString(properties, "style:text-underline-style") === "solid"
          ? "single"
          : undefined,
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
    });
  }
  return result;
}

function parseParagraph(
  paragraph: Element,
  textStyles: Map<string, TextProperties>,
): ParagraphDescriptorOptions {
  const children: (string | TextRunOptions)[] = [];
  for (const child of paragraph.elements ?? []) {
    if (child.type === "text") children.push(String(child.text ?? ""));
    else if (child.name === "text:span") {
      children.push({
        text: textOf(child),
        ...textStyles.get(attributeString(child, "text:style-name") ?? ""),
      });
    }
  }
  if (children.length === 1 && typeof children[0] === "string") return { text: children[0] };
  return children.length ? { children } : { text: "" };
}

import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  graphicFill,
  graphicOutline,
  lengthToEmu,
  PRESET_GEOMETRY_DOCX,
  presetGeometryOdf,
  pushShapeStyle,
  xmlElement,
  type GraphicStyle,
} from "@office-open/odf";
import type {
  ConnectorOptions,
  LineShapeOptions,
  ShapeOptions,
  SlideChild,
} from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { endpointConnection, parseChartFrame, parsePictureFrame } from "./drawing";
import { unknownSlideChild } from "./presentation";
import { parseSlideTable } from "./slide";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";

/** Shape with preset geometry renders as draw:custom-shape + enhanced-geometry. */
export function customShapeXml(shape: ShapeOptions, styles: string[]): string {
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

export function shapeXml(shape: ShapeOptions, styles: string[]): string {
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

/** Serializes a connector as a straight draw:connector with endpoint gluing. */
export function connectorXml(connector: ConnectorOptions): string {
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
export function lineXml(line: LineShapeOptions): string {
  return xmlElement("draw:line", {
    "draw:name": line.name,
    "svg:x1": toOdfLength(line.x1),
    "svg:y1": toOdfLength(line.y1),
    "svg:x2": toOdfLength(line.x2),
    "svg:y2": toOdfLength(line.y2),
  });
}

/** pptx EMU numbers pass through; universal measures resolve to EMU first. */
export function toOdfLength(value: number | string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return emuToLength(typeof value === "number" ? value : (lengthToEmu(value) ?? 0));
}

export function parseShape(frame: Element, textStyles: Map<string, TextProperties>): ShapeOptions {
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
export function parseCustomSlideShape(
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
export function parseConnector(element: Element): { connector: ConnectorOptions } {
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
export function parseLine(element: Element): { line: LineShapeOptions } {
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
export function parseGroup(
  element: Element,
  textStyles: Map<string, TextProperties>,
  columnWidths: Map<string, number>,
  binaries: Record<string, Uint8Array>,
  chartPool: Map<string, ChartSpaceOptions>,
  usedBinaryPaths: Set<string>,
): SlideChild[] {
  const children = (element.elements ?? []).flatMap((child): SlideChild[] => {
    if (child.name === "draw:frame") {
      const chart = parseChartFrame(child, chartPool);
      if (chart) return [chart];
      if (childNamed(child, "draw:object-ole")) return [];
      const picture = parsePictureFrame(child, binaries, usedBinaryPaths);
      if (picture) return [picture];
      if (childNamed(child, "draw:text-box")) return [{ shape: parseShape(child, textStyles) }];
      throw unknownSlideChild(child, "/draw:g");
    }
    if (child.name === "draw:line") return [parseLine(child)];
    if (child.name === "draw:connector") return [parseConnector(child)];
    if (child.name === "draw:g")
      return parseGroup(child, textStyles, columnWidths, binaries, chartPool, usedBinaryPaths);
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

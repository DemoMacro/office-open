import type { ChartSpaceOptions, CustomGeometryOptions, PathCommand } from "@office-open/core";
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
import { stringifyElement } from "@office-open/xml";

import { endpointConnection, parseChartFrame, parsePictureFrame } from "./drawing";
import { OdpParseError } from "./error";
import { unknownSlideChild } from "./presentation";
import { parseSlideTable } from "./slide";
import { parseParagraph, textBodyXml, type TextProperties } from "./text-run";

/** Shape with preset geometry renders as draw:custom-shape + enhanced-geometry. */
export function odfShapeId(element: Element): number | undefined {
  const match = /^id(\d+)$/u.exec(attributeString(element, "draw:id") ?? "");
  const id = match === null ? undefined : Number(match[1]);
  return id !== undefined && id > 0 ? id : undefined;
}

export function odfShapeIdXml(id: number | undefined): string | undefined {
  return id === undefined ? undefined : `id${id}`;
}

/** Collect every ODF shape identifier and project it onto PPTX numeric ids. */
export function collectOdfShapeIds(root: Element): Map<string, number> {
  const rawIds: string[] = [];
  const textAliases: [string, string | undefined][] = [];
  const visit = (element: Element, shapeId: string | undefined): void => {
    const drawId = attributeString(element, "draw:id");
    const currentShapeId = isShapeElement(element) ? drawId : shapeId;
    if (drawId !== undefined) rawIds.push(drawId);
    const textId = attributeString(element, "text:id");
    if (textId !== undefined) {
      rawIds.push(textId);
      textAliases.push([textId, currentShapeId]);
    }
    for (const child of element.elements ?? [])
      if (child.type === "element") visit(child, currentShapeId);
  };
  visit(root, undefined);

  let nextId = 1;
  const references = new Map<string, number>();
  for (const raw of rawIds) {
    const numbered = /^id(\d+)$/u.exec(raw);
    const id = numbered === null ? undefined : Number(numbered[1]);
    if (id !== undefined && id > 0) {
      references.set(raw, id);
      nextId = Math.max(nextId, id + 1);
    }
  }
  for (const raw of rawIds) {
    if (references.has(raw)) continue;
    references.set(raw, nextId++);
  }
  for (const [textId, shapeId] of textAliases) {
    const resolved = shapeId === undefined ? undefined : references.get(shapeId);
    if (resolved !== undefined) references.set(textId, resolved);
  }
  return references;
}

function isShapeElement(element: Element): boolean {
  const [namespace, localName] = element.name?.split(":", 2) ?? [];
  return (
    namespace === "draw" &&
    [
      "circle",
      "connector",
      "control",
      "custom-shape",
      "ellipse",
      "frame",
      "g",
      "line",
      "path",
      "polygon",
      "polyline",
      "rect",
    ].includes(localName ?? "")
  );
}

function shapeReferenceKey(element: Element): string | undefined {
  return attributeString(element, "draw:id") ?? descendantTextId(element);
}

function descendantTextId(element: Element): string | undefined {
  const direct = attributeString(element, "text:id");
  if (direct !== undefined) return direct;
  for (const child of element.elements ?? []) {
    const id = descendantTextId(child);
    if (id !== undefined) return id;
  }
  return undefined;
}

/** Assign parsed ODF shape identifiers to the matching canonical shape union. */
export function assignOdfShapeIds(
  element: Element,
  children: readonly SlideChild[],
  references: Map<string, number>,
): void {
  const assignShape = (child: SlideChild): void => {
    if (!("shape" in child)) return;
    const id = references.get(shapeReferenceKey(element) ?? "");
    if (id !== undefined) child.shape.id = id;
  };
  if (element.name === "draw:g") {
    const group = children.find(
      (child): child is Extract<SlideChild, { group: object }> => "group" in child,
    )?.group;
    if (group) {
      const id = references.get(shapeReferenceKey(element) ?? "");
      if (id !== undefined) group.id = id;
      const sourceChildren = (element.elements ?? []).filter(
        (child): child is Element => child.type === "element",
      );
      let childIndex = 0;
      for (const source of sourceChildren) {
        if (
          source.name !== "presentation:notes" &&
          source.name !== "office:forms" &&
          source.name !== "officeooo:annotation" &&
          source.name !== undefined &&
          !isAnimationElementName(source.name)
        ) {
          assignOdfShapeIds(source, group.children.slice(childIndex), references);
          childIndex++;
        }
      }
    }
    return;
  }
  assignShape((children[0] ?? {}) as SlideChild);
}

function isAnimationElementName(name: string): boolean {
  return name.startsWith("anim:") || name === "loext:animatePhysics";
}

export function customShapeXml(shape: ShapeOptions, styles: string[]): string {
  if (shape.controlId !== undefined) return controlShapeXml(shape, styles);
  const geometry = shape.properties?.geometry;
  const preset = typeof geometry === "string" ? geometry : geometry?.preset;
  const customGeometry = shape.properties?.customGeometry;
  if (customGeometry)
    return xmlElement(
      "draw:path",
      {
        "draw:id": odfShapeIdXml(shape.id),
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
        "svg:d": customGeometryPath(customGeometry),
      },
      shape.textBody ? textBodyXml(shape.textBody, styles) : [],
    );
  return xmlElement(
    "draw:custom-shape",
    {
      "draw:id": odfShapeIdXml(shape.id),
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
      ...(shape.textBody ? textBodyXml(shape.textBody, styles) : []),
    ],
  );
}

/** ODF form-control placeholder restores as a positioned draw:control. */
export function controlShapeXml(shape: ShapeOptions, styles: string[]): string {
  return xmlElement(
    "draw:control",
    {
      "draw:control": shape.controlId,
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
    shape.textBody ? textBodyXml(shape.textBody, styles) : [],
  );
}

function customGeometryPath(geometry: CustomGeometryOptions): string {
  return (geometry.pathList ?? [])
    .flatMap((path) => path.commands ?? [])
    .map((command): string => {
      const coordinate = (value: string): string =>
        String(Math.round((Number(value) / 360000) * 1000) / 1000);
      if (command.command === "moveTo")
        return `M ${coordinate(command.point.x)} ${coordinate(command.point.y)}`;
      if (command.command === "lineTo")
        return `L ${coordinate(command.point.x)} ${coordinate(command.point.y)}`;
      if (command.command === "close") return "Z";
      throw new OdpParseError(
        "Unsupported custom geometry path command",
        "content.xml",
        "/draw:page/draw:path",
        "svg:d",
        "unsupported-path-command",
      );
    })
    .join(" ");
}

export function shapeXml(shape: ShapeOptions, styles: string[]): string {
  return xmlElement(
    "draw:frame",
    {
      "draw:id": odfShapeIdXml(shape.id),
      "draw:name": shape.name,
      "svg:x": emuToLength(shape.x),
      "svg:y": emuToLength(shape.y),
      "svg:width": emuToLength(shape.width),
      "svg:height": emuToLength(shape.height),
    },
    [
      xmlElement(
        "draw:text-box",
        undefined,
        shape.textBody ? textBodyXml(shape.textBody, styles) : [],
      ),
    ],
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

export function parseShape(
  frame: Element,
  textStyles: Map<string, TextProperties>,
  id?: number,
): ShapeOptions {
  const textBox = childNamed(frame, "draw:text-box");
  const paragraphs = childrenNamed(textBox, "text:p");
  const shapeId = id ?? odfShapeId(frame);
  return {
    name: attributeString(frame, "draw:name"),
    x: lengthToEmu(attributeString(frame, "svg:x")),
    y: lengthToEmu(attributeString(frame, "svg:y")),
    width: lengthToEmu(attributeString(frame, "svg:width")),
    height: lengthToEmu(attributeString(frame, "svg:height")),
    ...(shapeId === undefined ? {} : { id: shapeId }),
    ...(paragraphs.length > 0
      ? {
          textBody: {
            paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)),
          },
        }
      : {}),
  };
}

/** draw:control maps to a typed pptx shape with its ODF form-control overlay. */
export function parseControlSlideShape(
  element: Element,
  textStyles: Map<string, TextProperties>,
  graphicStyles: Map<string, GraphicStyle>,
): ShapeOptions {
  const graphic = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const paragraphs = childrenNamed(element, "text:p");
  return {
    controlId: attributeString(element, "draw:control"),
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
    name: attributeString(element, "draw:name"),
    x: lengthToEmu(attributeString(element, "svg:x")),
    y: lengthToEmu(attributeString(element, "svg:y")),
    width: lengthToEmu(attributeString(element, "svg:width")),
    height: lengthToEmu(attributeString(element, "svg:height")),
    properties: {
      geometry: "rect",
      fill: graphicFill(graphic),
      outline: graphicOutline(graphic),
    },
    ...(paragraphs.length > 0
      ? {
          textBody: {
            paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)),
          },
        }
      : {}),
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
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
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

/** ODF primitive/path shapes map onto the shared pptx shape model. */
export function parsePrimitiveSlideShape(
  element: Element,
  textStyles: Map<string, TextProperties>,
  graphicStyles: Map<string, GraphicStyle>,
): ShapeOptions {
  const graphic = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const fill = graphicFill(graphic);
  const outline = graphicOutline(graphic);
  const paragraphs = childrenNamed(element, "text:p");
  const x = lengthToEmu(attributeString(element, "svg:x"));
  const y = lengthToEmu(attributeString(element, "svg:y"));
  const customGeometry = primitiveGeometry(element);
  const geometry =
    element.name === "draw:rect"
      ? ("rect" as const)
      : element.name === "draw:ellipse" || element.name === "draw:circle"
        ? ("ellipse" as const)
        : undefined;
  return {
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
    name: attributeString(element, "draw:name"),
    x,
    y,
    width: lengthToEmu(attributeString(element, "svg:width")),
    height: lengthToEmu(attributeString(element, "svg:height")),
    ...(paragraphs.length > 0
      ? {
          textBody: {
            paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)),
          },
        }
      : {}),
    ...(geometry || fill || outline || customGeometry
      ? {
          properties: {
            ...(geometry ? { geometry } : {}),
            ...(customGeometry ? { customGeometry } : {}),
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
  graphicStyles: Map<string, GraphicStyle>,
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
      if (child.name === "draw:frame") return [{ rawXml: stringifyElement(child) }];
      throw unknownSlideChild(child, "/draw:g");
    }
    if (child.name === "draw:custom-shape")
      return [{ shape: parseCustomSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:control")
      return [{ shape: parseControlSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:line") return [parseLine(child)];
    if (isPrimitiveShape(child))
      return [{ shape: parsePrimitiveSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:connector") return [parseConnector(child)];
    if (child.name === "draw:g")
      return parseGroup(
        child,
        textStyles,
        graphicStyles,
        columnWidths,
        binaries,
        chartPool,
        usedBinaryPaths,
      );
    if (child.name === "table:table") return [parseSlideTable(child, textStyles, columnWidths)];
    if (child.name === "dr3d:scene") return [{ rawXml: stringifyElement(child) }];
    throw unknownSlideChild(child, "/draw:g");
  });
  return [
    {
      group: {
        ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
        name: attributeString(element, "draw:name"),
        children,
      },
    },
  ];
}

function isPrimitiveShape(element: Element): boolean {
  return (
    element.name === "draw:rect" ||
    element.name === "draw:circle" ||
    element.name === "draw:ellipse" ||
    element.name === "draw:path" ||
    element.name === "draw:polygon" ||
    element.name === "draw:polyline"
  );
}

function primitiveGeometry(element: Element): CustomGeometryOptions | undefined {
  if (element.name === "draw:path")
    return {
      pathList: [{ commands: parseSvgPath(attributeString(element, "svg:d") ?? "") }],
    };
  if (element.name === "draw:polygon" || element.name === "draw:polyline")
    return {
      pathList: [
        {
          commands: parsePointList(
            attributeString(element, "draw:points") ?? attributeString(element, "svg:points") ?? "",
            element.name === "draw:polygon",
          ),
        },
      ],
    };
  return undefined;
}

function parseSvgPath(path: string): PathCommand[] {
  const tokens = path.match(/[MLHVCSQTAZmlhvcsqtaz]|-?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/gu);
  if (!tokens) throw invalidPrimitivePath();
  const commands: PathCommand[] = [];
  let index = 0;
  let command:
    | "M"
    | "L"
    | "H"
    | "V"
    | "C"
    | "S"
    | "Q"
    | "T"
    | "A"
    | "Z"
    | "m"
    | "l"
    | "h"
    | "v"
    | "c"
    | "s"
    | "q"
    | "t"
    | "a"
    | "z"
    | undefined;
  let currentX = 0;
  let currentY = 0;
  let startX = 0;
  let startY = 0;
  let previousCubic: { x: number; y: number } | undefined;
  let previousQuad: { x: number; y: number } | undefined;
  const point = (x: number, y: number, relative: boolean): { x: string; y: string } => {
    const left = relative ? currentX + x : x;
    const top = relative ? currentY + y : y;
    return {
      x: String(lengthToEmu(`${left}cm`) ?? 0),
      y: String(lengthToEmu(`${top}cm`) ?? 0),
    };
  };
  while (index < tokens.length) {
    const token = tokens[index]!;
    if (/^[MLHVCSQTAZmlhvcsqtaz]$/u.test(token)) {
      if (token === "Z" || token === "z") {
        commands.push({ command: "close" });
        currentX = startX;
        currentY = startY;
        command = undefined;
      } else {
        command = token as typeof command;
      }
      index++;
      continue;
    }
    if (!command) throw invalidPrimitivePath();
    const relative = command === command.toLowerCase();
    const argumentCounts = { M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7 };
    const count = argumentCounts[command.toUpperCase() as keyof typeof argumentCounts] ?? 0;
    if (count === 0) throw invalidPrimitivePath();
    const values: number[] = [];
    for (let position = 0; position < count; position++) {
      const value = Number(tokens[index]);
      if (!Number.isFinite(value)) throw invalidPrimitivePath();
      values.push(value);
      index++;
    }
    const upper = command.toUpperCase() as "M" | "L" | "H" | "V" | "C" | "S" | "Q" | "T" | "A";
    if (upper === "M") {
      const [x, y] = [values[0]!, values[1]!];
      const target = point(x, y, relative);
      currentX = relative ? currentX + x : x;
      currentY = relative ? currentY + y : y;
      startX = currentX;
      startY = currentY;
      commands.push({ command: "moveTo", point: target });
      command = relative ? "l" : "L";
    } else if (upper === "L" || upper === "H" || upper === "V") {
      const x = upper === "V" ? currentX : relative ? currentX + values[0]! : values[0]!;
      const y =
        upper === "H"
          ? currentY
          : relative
            ? currentY + (upper === "V" ? values[0]! : values[1]!)
            : upper === "V"
              ? values[0]!
              : values[1]!;
      currentX = x;
      currentY = y;
      commands.push({ command: "lineTo", point: point(x, y, false) });
    } else if (upper === "C" || upper === "S") {
      const first =
        upper === "C"
          ? {
              x: relative ? currentX + values[0]! : values[0]!,
              y: relative ? currentY + values[1]! : values[1]!,
            }
          : {
              x: previousCubic ? 2 * currentX - previousCubic.x : currentX,
              y: previousCubic ? 2 * currentY - previousCubic.y : currentY,
            };
      const second = {
        x: relative ? currentX + values[upper === "C" ? 2 : 0]! : values[upper === "C" ? 2 : 0]!,
        y: relative ? currentY + values[upper === "C" ? 3 : 1]! : values[upper === "C" ? 3 : 1]!,
      };
      const end = {
        x: relative ? currentX + values[upper === "C" ? 4 : 2]! : values[upper === "C" ? 4 : 2]!,
        y: relative ? currentY + values[upper === "C" ? 5 : 3]! : values[upper === "C" ? 5 : 3]!,
      };
      commands.push({
        command: "cubicBezTo",
        points: [
          point(first.x, first.y, false),
          point(second.x, second.y, false),
          point(end.x, end.y, false),
        ],
      });
      previousCubic = second;
      previousQuad = undefined;
      currentX = end.x;
      currentY = end.y;
    } else if (upper === "Q" || upper === "T") {
      const control =
        upper === "Q"
          ? {
              x: relative ? currentX + values[0]! : values[0]!,
              y: relative ? currentY + values[1]! : values[1]!,
            }
          : {
              x: previousQuad ? 2 * currentX - previousQuad.x : currentX,
              y: previousQuad ? 2 * currentY - previousQuad.y : currentY,
            };
      const end = {
        x: relative ? currentX + values[upper === "Q" ? 2 : 0]! : values[upper === "Q" ? 2 : 0]!,
        y: relative ? currentY + values[upper === "Q" ? 3 : 1]! : values[upper === "Q" ? 3 : 1]!,
      };
      commands.push({
        command: "quadBezTo",
        points: [point(control.x, control.y, false), point(end.x, end.y, false)],
      });
      previousQuad = control;
      previousCubic = undefined;
      currentX = end.x;
      currentY = end.y;
    } else {
      const [radiusX, radiusY, rotation, large, sweep, x, y] = values;
      const end = { x: relative ? currentX + x! : x!, y: relative ? currentY + y! : y! };
      commands.push(
        ...arcToCubic(
          currentX,
          currentY,
          radiusX!,
          radiusY!,
          rotation!,
          large! !== 0,
          sweep! !== 0,
          end.x,
          end.y,
        ),
      );
      previousCubic = undefined;
      previousQuad = undefined;
      currentX = end.x;
      currentY = end.y;
    }
  }
  if (commands.length === 0) throw invalidPrimitivePath();
  return commands;
}

function arcToCubic(
  startX: number,
  startY: number,
  radiusX: number,
  radiusY: number,
  rotationDegrees: number,
  largeArc: boolean,
  sweep: boolean,
  endX: number,
  endY: number,
): PathCommand[] {
  if (startX === endX && startY === endY) return [];
  const phi = (rotationDegrees * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  let rx = Math.abs(radiusX);
  let ry = Math.abs(radiusY);
  const left = (cos * (startX - endX)) / 2 + (sin * (startY - endY)) / 2;
  const top = (-sin * (startX - endX)) / 2 + (cos * (startY - endY)) / 2;
  const scale = (left * left) / (rx * rx) + (top * top) / (ry * ry);
  if (scale > 1) {
    const factor = Math.sqrt(scale);
    rx *= factor;
    ry *= factor;
  }
  const denominator = rx * rx * top * top + ry * ry * left * left;
  const factor = Math.sqrt(Math.max(0, (rx * rx * ry * ry - denominator) / denominator));
  const sign = largeArc === sweep ? -1 : 1;
  const centerLeft = (sign * factor * rx * top) / ry;
  const centerTop = (-sign * factor * ry * left) / rx;
  const centerX = cos * centerLeft - sin * centerTop + (startX + endX) / 2;
  const centerY = sin * centerLeft + cos * centerTop + (startY + endY) / 2;
  const angle = (x: number, y: number): number =>
    Math.atan2((y - centerY) / ry, (x - centerX) / rx);
  let theta = angle(startX, startY);
  let delta = angle(endX, endY) - theta;
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.ceil(Math.abs(delta) / (Math.PI / 2));
  const commands: PathCommand[] = [];
  for (let segment = 1; segment <= segments; segment++) {
    const start = theta + (delta * (segment - 1)) / segments;
    const end = theta + (delta * segment) / segments;
    const alpha = (4 / 3) * Math.tan((end - start) / 4);
    const local = (angleValue: number, radiusFactor: number): { x: number; y: number } => ({
      x: centerX + rx * Math.cos(angleValue) - alpha * radiusFactor * rx * Math.sin(angleValue),
      y: centerY + ry * Math.sin(angleValue) + alpha * radiusFactor * ry * Math.cos(angleValue),
    });
    const first = local(start, 1);
    const second = local(end, -1);
    const target = segment === segments ? { x: endX, y: endY } : local(end, 0);
    commands.push({
      command: "cubicBezTo",
      points: [
        {
          x: String(lengthToEmu(`${first.x}cm`) ?? 0),
          y: String(lengthToEmu(`${first.y}cm`) ?? 0),
        },
        {
          x: String(lengthToEmu(`${second.x}cm`) ?? 0),
          y: String(lengthToEmu(`${second.y}cm`) ?? 0),
        },
        {
          x: String(lengthToEmu(`${target.x}cm`) ?? 0),
          y: String(lengthToEmu(`${target.y}cm`) ?? 0),
        },
      ],
    });
  }
  return commands;
}

function parsePointList(points: string, close: boolean): PathCommand[] {
  const values = points
    .split(/[ ,;]+/u)
    .filter(Boolean)
    .map(Number);
  if (
    values.length < 4 ||
    values.length % 2 !== 0 ||
    values.some((value) => !Number.isFinite(value))
  )
    throw invalidPrimitivePath();
  const commands: PathCommand[] = [];
  for (let index = 0; index < values.length; index += 2) {
    commands.push({
      command: index === 0 ? "moveTo" : "lineTo",
      point: {
        x: String(lengthToEmu(`${values[index]!}cm`) ?? 0),
        y: String(lengthToEmu(`${values[index + 1]!}cm`) ?? 0),
      },
    });
  }
  return close ? [...commands, { command: "close" }] : commands;
}

function invalidPrimitivePath(): OdpParseError {
  return new OdpParseError(
    "Invalid ODF primitive shape geometry",
    "content.xml",
    "/draw:page",
    "svg:points",
    "invalid-primitive-geometry",
  );
}

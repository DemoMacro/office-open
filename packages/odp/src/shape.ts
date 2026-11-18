import type { ChartSpaceOptions, CustomGeometryOptions } from "@office-open/core";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  graphicFill,
  graphicOutline,
  lengthToEmu,
  odfEnhancedGeometryXml,
  odfGroupTransform,
  parseOdfEnhancedGeometry,
  parseOdfGroupTransform,
  parseOdfPrimitiveGeometry,
  presetGeometryOdf,
  pushShapeStyle,
  resolveOdfDrawGeometry,
  resolveOdfDrawGeometryOverlay,
  xmlElement,
  textOf,
  type GraphicStyle,
} from "@office-open/odf";
import type {
  ConnectorOptions,
  GroupOptions,
  LineShapeOptions,
  ShapeOptions,
  SlideChild,
} from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { endpointConnection, parseChartFrame, parsePictureFrame } from "./drawing";
import { OdpParseError, unknownSlideChild } from "./error";
import { parseSlideTable } from "./slide";
import { parseSlideNode, type OdpSlideNodeOptions } from "./slide-nodes";
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
  const shapeNames: string[] = [];
  const collectShapeNames = (element: Element): void => {
    const name = isShapeElement(element) ? attributeString(element, "draw:name") : undefined;
    if (name !== undefined) shapeNames.push(name);
    for (const child of element.elements ?? [])
      if (child.type === "element") collectShapeNames(child);
  };
  collectShapeNames(root);
  for (const name of shapeNames) {
    if (!references.has(name)) references.set(name, nextId++);
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
        "draw:transform": odfGroupTransform(shape.rotation),
        "draw:name": shape.name,
      },
      [
        odfEnhancedGeometryXml(customGeometry, shape.odfGeometryType),
        ...(shape.textBody ? textBodyXml(shape.textBody, styles) : []),
      ],
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
      "draw:transform": odfGroupTransform(shape.rotation),
      "draw:name": shape.name,
    },
    [
      xmlElement(
        "draw:enhanced-geometry",
        {
          "draw:type": shape.odfGeometryType ?? (preset ? presetGeometryOdf(preset) : undefined),
        },
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
      "draw:transform": odfGroupTransform(shape.rotation),
      "draw:name": shape.name,
    },
    shape.textBody ? textBodyXml(shape.textBody, styles) : [],
  );
}

export function shapeXml(shape: ShapeOptions, styles: string[]): string {
  return xmlElement(
    "draw:frame",
    {
      "draw:id": odfShapeIdXml(shape.id),
      "draw:name": shape.name,
      "draw:style-name": pushShapeStyle(shape.properties?.fill, shape.properties?.outline, styles),
      "svg:x": emuToLength(shape.x),
      "svg:y": emuToLength(shape.y),
      "svg:width": emuToLength(shape.width),
      "svg:height": emuToLength(shape.height),
      "draw:transform": odfGroupTransform(shape.rotation),
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
export function connectorXml(connector: ConnectorOptions, styles: string[]): string {
  return xmlElement("draw:connector", {
    "draw:id": odfShapeIdXml(connector.id),
    "draw:type": "line",
    "draw:name": connector.name,
    "draw:style-name": pushShapeStyle(
      connector.properties?.fill,
      connector.properties?.outline,
      styles,
    ),
    "svg:x1": toOdfLength(connector.x1),
    "svg:y1": toOdfLength(connector.y1),
    "svg:x2": toOdfLength(connector.x2),
    "svg:y2": toOdfLength(connector.y2),
    "draw:transform": odfGroupTransform(connector.rotation),
    "draw:start-shape": connector.startConnection
      ? String(connector.startConnection.id)
      : undefined,
    "draw:start-glue-point": connector.startConnection?.index,
    "draw:end-shape": connector.endConnection ? String(connector.endConnection.id) : undefined,
    "draw:end-glue-point": connector.endConnection?.index,
  });
}

/** Serializes a straight line as draw:line. */
export function lineXml(line: LineShapeOptions, styles: string[]): string {
  return xmlElement("draw:line", {
    "draw:id": odfShapeIdXml(line.id),
    "draw:name": line.name,
    "draw:style-name": pushShapeStyle(line.properties?.fill, line.properties?.outline, styles),
    "svg:x1": toOdfLength(line.x1),
    "svg:y1": toOdfLength(line.y1),
    "svg:x2": toOdfLength(line.x2),
    "svg:y2": toOdfLength(line.y2),
    "draw:transform": odfGroupTransform(line.rotation),
  });
}

/** pptx EMU numbers pass through; universal measures resolve to EMU first. */
export function toOdfLength(value: number | string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return emuToLength(typeof value === "number" ? value : (lengthToEmu(value) ?? 0));
}

function parseShapeTransform(element: Element): GroupTransform {
  const transform = parseOdfGroupTransform(attributeString(element, "draw:transform"));
  if (transform === null)
    throw new OdpParseError(
      `ODP ${element.name ?? "shape"} transform cannot be projected to the canonical shape model`,
      "content.xml",
      `/office:document-content/office:body/office:presentation/${element.name ?? "shape"}`,
      "draw:transform",
      "unsupported-shape-transform",
    );
  return transform;
}

function transformedLength(
  value: string | undefined,
  offset: number | undefined,
): number | undefined {
  return transformedLengthValue(lengthToEmu(value), offset);
}

function transformedLengthValue(
  value: number | undefined,
  offset: number | undefined,
): number | undefined {
  if (value === undefined) return offset;
  return value + (offset ?? 0);
}

export function parseShape(
  frame: Element,
  textStyles: Map<string, TextProperties>,
  graphicStyles: Map<string, GraphicStyle> = new Map(),
  id?: number,
): ShapeOptions {
  const textBox = childNamed(frame, "draw:text-box");
  const paragraphs = childrenNamed(textBox, "text:p");
  const shapeId = id ?? odfShapeId(frame);
  const transform = parseShapeTransform(frame);
  const graphic = graphicStyles.get(attributeString(frame, "draw:style-name") ?? "");
  const fill = graphicFill(graphic);
  const outline = graphicOutline(graphic);
  return {
    name: attributeString(frame, "draw:name"),
    x: transformedLength(attributeString(frame, "svg:x"), transform.x),
    y: transformedLength(attributeString(frame, "svg:y"), transform.y),
    width: lengthToEmu(attributeString(frame, "svg:width")),
    height: lengthToEmu(attributeString(frame, "svg:height")),
    ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
    ...(shapeId === undefined ? {} : { id: shapeId }),
    ...(fill || outline ? { properties: { fill, outline } } : {}),
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
  const transform = parseShapeTransform(element);
  const paragraphs = childrenNamed(element, "text:p");
  return {
    controlId: attributeString(element, "draw:control"),
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
    name: attributeString(element, "draw:name"),
    x: transformedLength(attributeString(element, "svg:x"), transform.x),
    y: transformedLength(attributeString(element, "svg:y"), transform.y),
    width: lengthToEmu(attributeString(element, "svg:width")),
    height: lengthToEmu(attributeString(element, "svg:height")),
    ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
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
  const resolvedGeometry = presetType ? resolveOdfDrawGeometry(presetType) : undefined;
  const preset = resolvedGeometry?.kind === "shape" ? resolvedGeometry.preset : undefined;
  const textWarp = resolvedGeometry?.kind === "text-warp" ? resolvedGeometry.preset : undefined;
  const geometryOverlay =
    presetType && resolvedGeometry
      ? resolveOdfDrawGeometryOverlay(presetType, resolvedGeometry)
      : undefined;
  const customGeometry = parseOdfEnhancedGeometry(enhanced);
  const paragraphs = childrenNamed(element, "text:p");
  const transform = parseShapeTransform(element);
  const fill = graphicFill(graphic);
  const outline = graphicOutline(graphic);
  return {
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
    name: attributeString(element, "draw:name"),
    x: transformedLength(attributeString(element, "svg:x"), transform.x),
    y: transformedLength(attributeString(element, "svg:y"), transform.y),
    width: lengthToEmu(attributeString(element, "svg:width")),
    height: lengthToEmu(attributeString(element, "svg:height")),
    ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
    ...(paragraphs.length > 0
      ? {
          textBody: {
            paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)),
          },
        }
      : {}),
    ...(preset || fill || outline || customGeometry
      ? {
          properties: {
            ...(preset ? { geometry: preset } : {}),
            ...(fill ? { fill } : {}),
            ...(outline ? { outline } : {}),
            ...(customGeometry ? { customGeometry } : {}),
          },
        }
      : {}),
    ...(geometryOverlay ? { odfGeometryType: geometryOverlay } : {}),
    ...(textWarp
      ? {
          textBody: {
            bodyProperties: { prstTxWarp: { preset: textWarp } },
            ...(paragraphs.length > 0
              ? { paragraphs: paragraphs.map((paragraph) => parseParagraph(paragraph, textStyles)) }
              : {}),
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
  const width = lengthToEmu(attributeString(element, "svg:width"));
  const height = lengthToEmu(attributeString(element, "svg:height"));
  const transform = parseShapeTransform(element);
  const customGeometry = primitiveGeometry(element, width ?? 0, height ?? 0);
  const geometry =
    element.name === "draw:rect"
      ? ("rect" as const)
      : element.name === "draw:ellipse" || element.name === "draw:circle"
        ? ("ellipse" as const)
        : undefined;
  return {
    ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
    name: attributeString(element, "draw:name"),
    x: transformedLengthValue(x, transform.x),
    y: transformedLengthValue(y, transform.y),
    width,
    height,
    ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
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
export function parseConnector(
  element: Element,
  graphicStyles: Map<string, GraphicStyle>,
): { connector: ConnectorOptions } {
  const graphic = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const transform = parseShapeTransform(element);
  return {
    connector: {
      ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
      name: attributeString(element, "draw:name"),
      x1: transformedLength(attributeString(element, "svg:x1"), transform.x),
      y1: transformedLength(attributeString(element, "svg:y1"), transform.y),
      x2: transformedLength(attributeString(element, "svg:x2"), transform.x),
      y2: transformedLength(attributeString(element, "svg:y2"), transform.y),
      ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
      startConnection: endpointConnection(element, "start"),
      endConnection: endpointConnection(element, "end"),
      ...(graphicFill(graphic) || graphicOutline(graphic)
        ? { properties: { fill: graphicFill(graphic), outline: graphicOutline(graphic) } }
        : {}),
    },
  };
}

/** Parses a draw:line back to a pptx line child. */
export function parseLine(
  element: Element,
  graphicStyles: Map<string, GraphicStyle>,
): { line: LineShapeOptions } {
  const graphic = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const transform = parseShapeTransform(element);
  return {
    line: {
      ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
      name: attributeString(element, "draw:name"),
      x1: transformedLength(attributeString(element, "svg:x1"), transform.x),
      y1: transformedLength(attributeString(element, "svg:y1"), transform.y),
      x2: transformedLength(attributeString(element, "svg:x2"), transform.x),
      y2: transformedLength(attributeString(element, "svg:y2"), transform.y),
      ...(transform.rotation === undefined ? {} : { rotation: transform.rotation }),
      ...(graphicFill(graphic) || graphicOutline(graphic)
        ? { properties: { fill: graphicFill(graphic), outline: graphicOutline(graphic) } }
        : {}),
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
): { group: GroupOptions & { odpNodes?: OdpSlideNodeOptions[] } }[] {
  const odpNodes: OdpSlideNodeOptions[] = [];
  const groupStyle = graphicStyles.get(attributeString(element, "draw:style-name") ?? "");
  const groupX = lengthToEmu(attributeString(element, "svg:x"));
  const groupY = lengthToEmu(attributeString(element, "svg:y"));
  const groupWidth = lengthToEmu(attributeString(element, "svg:width"));
  const groupHeight = lengthToEmu(attributeString(element, "svg:height"));
  const groupTransform = parseGroupTransform(attributeString(element, "draw:transform"));
  if (graphicOutline(groupStyle))
    throw new OdpParseError(
      "ODP group outline has no equivalent in the canonical group model",
      "content.xml",
      "/office:document-content/office:body/office:presentation/draw:page/draw:g",
      "draw:style-name",
      "unsupported-group-outline",
    );
  const children = (element.elements ?? []).flatMap((child): SlideChild[] => {
    if (child.name === "svg:title" || child.name === "svg:desc") return [];
    if (child.name === "draw:frame") {
      const chart = parseChartFrame(child, chartPool);
      if (chart) return [chart];
      const contentNames = (child.elements ?? [])
        .filter((candidate): candidate is Element => candidate.type === "element")
        .map((candidate) => candidate.name)
        .filter((name): name is string => name !== undefined);
      if (
        contentNames.includes("draw:object-ole") ||
        (contentNames.includes("draw:object") &&
          (childNamed(child, "svg:title") !== undefined ||
            childNamed(child, "svg:desc") !== undefined)) ||
        childNamed(child, "office:event-listeners") !== undefined ||
        (contentNames.filter((name) =>
          ["draw:object", "draw:image", "draw:plugin", "draw:text-box", "table:table"].includes(
            name,
          ),
        ).length > 1 &&
          !(
            (contentNames.includes("draw:object") &&
              contentNames.includes("draw:image") &&
              contentNames.length === 2) ||
            (contentNames.includes("draw:plugin") &&
              contentNames.includes("draw:image") &&
              contentNames.length === 2)
          ))
      ) {
        odpNodes.push({
          position: visibleChildIndex(element, child),
          node: parseSlideNode(child),
        });
        const node = odpNodes.at(-1)?.node;
        const frameChart = parseChartFrame(child, chartPool);
        if (node && "frame" in node && frameChart) {
          node.frame.contents = [{ kind: "chart", chart: frameChart.chart }];
        }
        return [];
      }
      const picture = parsePictureFrame(child, binaries, usedBinaryPaths, graphicStyles);
      if (picture) return [picture];
      if (childNamed(child, "draw:text-box"))
        return [{ shape: parseShape(child, textStyles, graphicStyles) }];
      odpNodes.push({
        position: visibleChildIndex(element, child),
        node: parseSlideNode(child),
      });
      return [];
    }
    if (child.name === "draw:custom-shape")
      return [{ shape: parseCustomSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:control")
      return [{ shape: parseControlSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:line") return [parseLine(child, graphicStyles)];
    if (isPrimitiveShape(child))
      return [{ shape: parsePrimitiveSlideShape(child, textStyles, graphicStyles) }];
    if (child.name === "draw:connector") return [parseConnector(child, graphicStyles)];
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
    if (child.name === "dr3d:scene") {
      odpNodes.push({
        position: visibleChildIndex(element, child),
        node: parseSlideNode(child, (frame) => parseChartFrame(frame, chartPool)),
      });
      return [];
    }
    throw unknownSlideChild(child, "/draw:g");
  });
  return [
    {
      group: {
        ...(odfShapeId(element) === undefined ? {} : { id: odfShapeId(element) }),
        name: attributeString(element, "draw:name"),
        ...(childNamed(element, "svg:title")
          ? { title: textOf(childNamed(element, "svg:title")) }
          : {}),
        ...(childNamed(element, "svg:desc")
          ? { description: textOf(childNamed(element, "svg:desc")) }
          : {}),
        children,
        ...(groupX === undefined &&
        groupTransform.x === undefined &&
        groupY === undefined &&
        groupTransform.y === undefined &&
        groupWidth === undefined &&
        groupHeight === undefined
          ? {}
          : {
              x: groupX === undefined ? groupTransform.x : groupX + (groupTransform.x ?? 0),
              y: groupY === undefined ? groupTransform.y : groupY + (groupTransform.y ?? 0),
              width: groupWidth,
              height: groupHeight,
            }),
        ...(groupTransform.rotation === undefined ? {} : { rotation: groupTransform.rotation }),
        ...(graphicFill(groupStyle) ? { fill: graphicFill(groupStyle) } : {}),
        ...(odpNodes.length > 0 ? { odpNodes } : {}),
      },
    },
  ];
}

interface GroupTransform {
  x?: number;
  y?: number;
  rotation?: number;
}

function parseGroupTransform(value: string | undefined): GroupTransform {
  const result = parseOdfGroupTransform(value);
  if (result === null) throw unsupportedGroupTransform(value ?? "");
  return result;
}

function unsupportedGroupTransform(value: string): never {
  throw new OdpParseError(
    `ODP group transform cannot be projected to the canonical group model: ${value}`,
    "content.xml",
    "/office:document-content/office:body/office:presentation/draw:g",
    "draw:transform",
    "unsupported-group-transform",
  );
}

function visibleChildIndex(container: Element, child: Element): number {
  let index = 0;
  for (const candidate of container.elements ?? []) {
    if (candidate.type !== "element") continue;
    if (candidate.name === "presentation:notes" || candidate.name === "office:forms") continue;
    if (candidate === child) return index;
    index++;
  }
  return index;
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

function primitiveGeometry(
  element: Element,
  widthEmu: number,
  heightEmu: number,
): CustomGeometryOptions | undefined {
  return parseOdfPrimitiveGeometry(element, widthEmu, heightEmu);
}

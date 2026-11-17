import type {
  AdjustHandle,
  AdjustPoint,
  ConnectionSite,
  CustomGeometryOptions,
  GeometryGuide,
  PathCommand,
  PathOptions,
} from "@office-open/core/drawing";
import type { Element } from "@office-open/xml";

import { OdfSchemaError } from "../errors";
import { attributeString, childrenNamed, xmlElement } from "../runtime/xml";

const ENHANCED_ATTRIBUTES = [
  "draw:type",
  "svg:viewBox",
  "draw:enhanced-path",
  "draw:text-areas",
  "draw:glue-points",
  "draw:modifiers",
] as const;

/**
 * Projects the ODF enhanced geometry common subset onto DrawingML custom geometry.
 * Unsupported ODF-only geometry controls fail explicitly instead of disappearing.
 */
export function parseOdfEnhancedGeometry(
  element: Element | undefined,
): CustomGeometryOptions | undefined {
  if (!element) return undefined;
  assertAttributes(element, ENHANCED_ATTRIBUTES, element);
  const unsupportedChild = (element.elements ?? []).find(
    (child) =>
      child.type === "element" && child.name !== "draw:equation" && child.name !== "draw:handle",
  );
  if (unsupportedChild?.type === "element") {
    throw geometryError(
      element,
      `child ${unsupportedChild.name ?? "unknown"} has no canonical customGeometry mapping`,
      unsupportedChild.name ?? "unknown",
    );
  }

  const modifiers = parseModifiers(attributeString(element, "draw:modifiers"));
  const equations = childrenNamed(element, "draw:equation").map(parseEquation);
  const path = attributeString(element, "draw:enhanced-path");
  const viewBox = parseViewBox(attributeString(element, "svg:viewBox"), element);
  const pathList = path === undefined ? undefined : parsePaths(path, element, viewBox);
  const textRectangle = parseTextRectangle(attributeString(element, "draw:text-areas"), element);
  const connectionSites = parseConnectionSites(
    attributeString(element, "draw:glue-points"),
    element,
  );
  const adjustHandles = childrenNamed(element, "draw:handle").map((handle) =>
    parseHandle(handle, element),
  );
  if (
    modifiers.length === 0 &&
    equations.length === 0 &&
    pathList === undefined &&
    textRectangle === undefined &&
    connectionSites === undefined &&
    adjustHandles.length === 0
  )
    return undefined;
  return {
    ...(modifiers.length > 0 ? { adjustmentValues: modifiers } : {}),
    ...(equations.length > 0 ? { guides: equations } : {}),
    ...(adjustHandles.length > 0 ? { adjustHandles } : {}),
    ...(connectionSites !== undefined ? { connectionSites } : {}),
    ...(textRectangle !== undefined ? { textRectangle } : {}),
    ...(pathList !== undefined ? { pathList } : {}),
  };
}

/** Serializes canonical custom geometry into an ODF enhanced-geometry element. */
export function odfEnhancedGeometryXml(geometry: CustomGeometryOptions, drawType?: string): string {
  const firstPath = geometry.pathList?.[0];
  const viewBox =
    firstPath?.w !== undefined && firstPath.h !== undefined
      ? `0 0 ${firstPath.w} ${firstPath.h}`
      : undefined;
  const modifiers = (geometry.adjustmentValues ?? [])
    .map((guide) => {
      const value = guide.formula.match(/^val\s+(.+)$/u)?.[1];
      if (value === undefined)
        throw generationError(
          `adjustment value ${guide.name} does not use a val formula`,
          "draw:modifiers",
        );
      return value;
    })
    .join(" ");
  const path = (geometry.pathList ?? [])
    .map((pathOptions) => pathOptions.commands?.map(odfPathCommand).join(" ") ?? "")
    .filter(Boolean)
    .join(" N ");
  const children = [
    ...(geometry.guides ?? []).map((guide) =>
      xmlElement("draw:equation", {
        "draw:name": guide.name,
        "draw:formula": guide.formula,
      }),
    ),
    ...(geometry.adjustHandles ?? []).map(odfHandleXml),
  ];
  return xmlElement(
    "draw:enhanced-geometry",
    {
      "draw:type": drawType,
      "svg:viewBox": viewBox,
      "draw:enhanced-path": path || undefined,
      "draw:text-areas": geometry.textRectangle
        ? [
            geometry.textRectangle.left,
            geometry.textRectangle.top,
            geometry.textRectangle.right,
            geometry.textRectangle.bottom,
          ].join(" ")
        : undefined,
      "draw:glue-points": geometry.connectionSites
        ?.map((site) => {
          if (site.angle !== "0")
            throw generationError(
              "connection site angles have no ODF enhanced-geometry equivalent",
              "draw:glue-points",
            );
          return `${odfToken(site.position.x)} ${odfToken(site.position.y)}`;
        })
        .join(" "),
      "draw:modifiers": modifiers || undefined,
    },
    children,
  );
}

function parseEquation(element: Element): GeometryGuide {
  const name = attributeString(element, "draw:name");
  const formula = attributeString(element, "draw:formula");
  if (name === undefined || formula === undefined)
    throw geometryError(
      element,
      "draw:equation requires draw:name and draw:formula",
      "draw:equation",
    );
  return { name, formula };
}

function parseModifiers(value: string | undefined): GeometryGuide[] {
  if (value === undefined || value.trim() === "") return [];
  return value
    .trim()
    .split(/\s+/u)
    .map((modifier, index) => ({ name: `modifier${index}`, formula: `val ${modifier}` }));
}

function parseViewBox(
  value: string | undefined,
  element: Element,
): readonly [number, number, number, number] | undefined {
  if (value === undefined) return undefined;
  const values = value
    .split(/[\s,]+/u)
    .filter(Boolean)
    .map(Number);
  if (values.length !== 4 || values.some((item) => !Number.isFinite(item)))
    throw geometryError(element, "svg:viewBox must contain four finite numbers", "svg:viewBox");
  if (values[0] !== 0 || values[1] !== 0)
    throw geometryError(
      element,
      "svg:viewBox with a non-zero origin has no canonical customGeometry mapping",
      "svg:viewBox",
    );
  return [values[0]!, values[1]!, values[2]!, values[3]!];
}

function parseTextRectangle(
  value: string | undefined,
  element: Element,
): CustomGeometryOptions["textRectangle"] {
  if (value === undefined) return undefined;
  const values = value.split(/[\s,]+/u).filter(Boolean);
  if (values.length !== 4)
    throw geometryError(element, "draw:text-areas must contain four values", "draw:text-areas");
  return { left: values[0]!, top: values[1]!, right: values[2]!, bottom: values[3]! };
}

function parseConnectionSites(
  value: string | undefined,
  element: Element,
): ConnectionSite[] | undefined {
  if (value === undefined) return undefined;
  const values = value.split(/[\s,]+/u).filter(Boolean);
  if (values.length % 2 !== 0)
    throw geometryError(
      element,
      "draw:glue-points must contain coordinate pairs",
      "draw:glue-points",
    );
  const sites: ConnectionSite[] = [];
  for (let index = 0; index < values.length; index += 2) {
    sites.push({
      angle: "0",
      position: { x: values[index]!, y: values[index + 1]! },
    });
  }
  return sites;
}

function parseHandle(handle: Element, parent: Element): AdjustHandle {
  assertAttributes(
    handle,
    [
      "draw:handle-position",
      "draw:handle-range-x-minimum",
      "draw:handle-range-x-maximum",
      "draw:handle-range-y-minimum",
      "draw:handle-range-y-maximum",
      "draw:handle-polar",
      "draw:handle-radius-range-minimum",
      "draw:handle-radius-range-maximum",
    ],
    parent,
  );
  if ((handle.elements ?? []).some((child) => child.type === "element"))
    throw geometryError(parent, "draw:handle must be empty", "draw:handle");
  const position = attributeString(handle, "draw:handle-position");
  if (position === undefined)
    throw geometryError(parent, "draw:handle-position is required", "draw:handle");
  const coordinates = position.split(/[\s,]+/u).filter(Boolean);
  if (coordinates.length !== 2)
    throw geometryError(
      parent,
      "draw:handle-position must contain two values",
      "draw:handle-position",
    );
  const polar = attributeString(handle, "draw:handle-polar");
  if (polar !== undefined) {
    if (
      attributeString(handle, "draw:handle-range-x-minimum") !== undefined ||
      attributeString(handle, "draw:handle-range-x-maximum") !== undefined ||
      attributeString(handle, "draw:handle-range-y-minimum") !== undefined ||
      attributeString(handle, "draw:handle-range-y-maximum") !== undefined
    )
      throw geometryError(parent, "polar draw:handle cannot also define XY ranges", "draw:handle");
    const references = polar.split(/[\s,]+/u).filter(Boolean);
    if (references.length !== 2)
      throw geometryError(parent, "draw:handle-polar must contain two values", "draw:handle-polar");
    return {
      type: "polar",
      guideRefRadius: references[0],
      guideRefAngle: references[1],
      minRadius: attributeString(handle, "draw:handle-radius-range-minimum"),
      maxRadius: attributeString(handle, "draw:handle-radius-range-maximum"),
      position: { x: coordinates[0]!, y: coordinates[1]! },
    };
  }
  return {
    type: "xy",
    minX: attributeString(handle, "draw:handle-range-x-minimum"),
    maxX: attributeString(handle, "draw:handle-range-x-maximum"),
    minY: attributeString(handle, "draw:handle-range-y-minimum"),
    maxY: attributeString(handle, "draw:handle-range-y-maximum"),
    position: { x: coordinates[0]!, y: coordinates[1]! },
  };
}

function parsePaths(
  source: string,
  element: Element,
  viewBox: readonly [number, number, number, number] | undefined,
): PathOptions[] {
  const tokens = source.split(/[\s,]+/u).filter(Boolean);
  const paths: PathOptions[] = [];
  let commands: PathCommand[] = [];
  const appendPath = (): void => {
    if (commands.length === 0) return;
    paths.push({
      ...(viewBox ? { w: viewBox[2], h: viewBox[3] } : {}),
      commands,
    });
    commands = [];
  };
  let index = 0;
  while (index < tokens.length) {
    const command = tokens[index++]?.toUpperCase();
    const count =
      command === "M" || command === "L"
        ? 2
        : command === "Q"
          ? 4
          : command === "C"
            ? 6
            : command === "A"
              ? 4
              : 0;
    if (command === "Z") {
      commands.push({ command: "close" });
      continue;
    }
    if (command === "N") {
      appendPath();
      continue;
    }
    if (count === 0)
      throw geometryError(
        element,
        `unsupported enhanced-path command ${command ?? ""}`,
        "draw:enhanced-path",
      );
    if (index + count > tokens.length)
      throw geometryError(
        element,
        "enhanced-path command has insufficient arguments",
        "draw:enhanced-path",
      );
    const args = tokens.slice(index, index + count).map(normalizeModifierToken);
    index += count;
    if (command === "M" || command === "L") {
      const point: AdjustPoint = { x: args[0]!, y: args[1]! };
      commands.push({ command: command === "M" ? "moveTo" : "lineTo", point });
    } else if (command === "Q") {
      commands.push({
        command: "quadBezTo",
        points: [
          { x: args[0]!, y: args[1]! },
          { x: args[2]!, y: args[3]! },
        ],
      });
    } else if (command === "C") {
      commands.push({
        command: "cubicBezTo",
        points: [
          { x: args[0]!, y: args[1]! },
          { x: args[2]!, y: args[3]! },
          { x: args[4]!, y: args[5]! },
        ],
      });
    } else {
      commands.push({
        command: "arcTo",
        widthRadius: args[0]!,
        heightRadius: args[1]!,
        startAngle: args[2]!,
        sweepAngle: args[3]!,
      });
    }
  }
  appendPath();
  return paths;
}

function normalizeModifierToken(value: string): string {
  const match = /^\$(\d+)$/u.exec(value);
  return match ? `modifier${match[1]}` : value;
}

function odfPathCommand(command: PathCommand): string {
  if (command.command === "close") return "Z";
  if (command.command === "moveTo")
    return `M ${odfToken(command.point.x)} ${odfToken(command.point.y)}`;
  if (command.command === "lineTo")
    return `L ${odfToken(command.point.x)} ${odfToken(command.point.y)}`;
  if (command.command === "quadBezTo")
    return `Q ${command.points.map((point) => `${odfToken(point.x)} ${odfToken(point.y)}`).join(" ")}`;
  if (command.command === "cubicBezTo")
    return `C ${command.points.map((point) => `${odfToken(point.x)} ${odfToken(point.y)}`).join(" ")}`;
  return `A ${odfToken(command.widthRadius)} ${odfToken(command.heightRadius)} ${odfToken(command.startAngle)} ${odfToken(command.sweepAngle)}`;
}

function odfToken(value: string): string {
  const match = /^modifier(\d+)$/u.exec(value);
  return match ? `$${match[1]}` : value;
}

function odfHandleXml(handle: AdjustHandle): string {
  const position = `${handle.position.x} ${handle.position.y}`;
  if (handle.type === "polar") {
    if (handle.minAngle !== undefined || handle.maxAngle !== undefined)
      throw generationError("polar handle angle ranges have no ODF equivalent", "draw:handle");
    return xmlElement("draw:handle", {
      "draw:handle-position": position,
      "draw:handle-polar": [handle.guideRefRadius, handle.guideRefAngle]
        .filter((value): value is string => value !== undefined)
        .map(odfToken)
        .join(" "),
      "draw:handle-radius-range-minimum": handle.minRadius,
      "draw:handle-radius-range-maximum": handle.maxRadius,
    });
  }
  if (handle.guideRefX !== undefined || handle.guideRefY !== undefined)
    throw generationError("XY handle guide references have no ODF equivalent", "draw:handle");
  return xmlElement("draw:handle", {
    "draw:handle-position": position,
    "draw:handle-range-x-minimum": handle.minX,
    "draw:handle-range-x-maximum": handle.maxX,
    "draw:handle-range-y-minimum": handle.minY,
    "draw:handle-range-y-maximum": handle.maxY,
  });
}

function assertAttributes(element: Element, allowed: readonly string[], parent: Element): void {
  const unsupported = Object.keys(element.attributes ?? {}).find((name) => !allowed.includes(name));
  if (unsupported !== undefined)
    throw geometryError(
      parent,
      `attribute ${unsupported} has no canonical customGeometry mapping`,
      unsupported,
    );
}

function generationError(reason: string, name: string): OdfSchemaError {
  return new OdfSchemaError(
    `ODF enhanced geometry: ${reason}`,
    "content.xml",
    "/draw:enhanced-geometry",
    name,
    reason,
  );
}

function geometryError(_element: Element, reason: string, name: string): OdfSchemaError {
  return new OdfSchemaError(
    `ODF enhanced geometry: ${reason}`,
    "content.xml",
    `/draw:enhanced-geometry/${name}`,
    name,
    reason,
  );
}

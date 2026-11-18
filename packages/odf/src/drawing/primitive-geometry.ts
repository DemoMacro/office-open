import type { CustomGeometryOptions, PathCommand } from "@office-open/core/drawing";
import type { Element } from "@office-open/xml";

import { OdfSchemaError } from "../errors";
import { attributeString } from "../runtime/xml";

export function parseOdfPrimitiveGeometry(
  element: Element,
  width: number,
  height: number,
): CustomGeometryOptions | undefined {
  if (
    element.name !== "draw:path" &&
    element.name !== "draw:polygon" &&
    element.name !== "draw:polyline"
  )
    return undefined;
  const localPoint = primitiveLocalPoint(element, width, height);
  const commands =
    element.name === "draw:path"
      ? parseSvgPath(attributeString(element, "svg:d") ?? "", localPoint)
      : parsePointList(
          attributeString(element, "draw:points") ?? attributeString(element, "svg:points") ?? "",
          element.name === "draw:polygon",
          localPoint,
        );
  return {
    pathList: [{ w: width, h: height, commands }],
  };
}

function primitiveLocalPoint(
  element: Element,
  width: number,
  height: number,
): (x: number, y: number) => { x: string; y: string } {
  const viewBox = parseViewBox(element, width, height);
  return (x, y) => ({
    x: String(Math.round(((x - viewBox.x) / viewBox.width) * width)),
    y: String(Math.round(((y - viewBox.y) / viewBox.height) * height)),
  });
}

function parseViewBox(
  element: Element,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } {
  const values = (attributeString(element, "svg:viewBox") ?? "")
    .split(/[\s,]+/u)
    .filter(Boolean)
    .map(Number);
  if (
    values.length !== 4 ||
    values.some((value) => !Number.isFinite(value)) ||
    values[2]! <= 0 ||
    values[3]! <= 0 ||
    width <= 0 ||
    height <= 0
  )
    throw primitiveGeometryError(
      element,
      "svg:viewBox is required and must be a finite rectangle",
      "svg:viewBox",
    );
  return { x: values[0]!, y: values[1]!, width: values[2]!, height: values[3]! };
}

function parseSvgPath(
  path: string,
  localPoint: (x: number, y: number) => { x: string; y: string },
): PathCommand[] {
  const tokens: string[] = [];
  const tokenPattern = /[MLHVCSQTAZmlhvcsqtaz]|[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/gu;
  let cursor = 0;
  while (cursor < path.length) {
    if (/[\s,]/u.test(path[cursor]!)) {
      cursor += 1;
      continue;
    }
    tokenPattern.lastIndex = cursor;
    const match = tokenPattern.exec(path);
    if (!match || match.index !== cursor)
      throw primitiveGeometryError(undefined, "invalid SVG path", "svg:d");
    tokens.push(match[0]);
    cursor += match[0].length;
  }
  const commands: PathCommand[] = [];
  let index = 0;
  let command: string | undefined;
  let currentX = 0;
  let currentY = 0;
  let startX = 0;
  let startY = 0;
  let previousCubic: { x: number; y: number } | undefined;
  let previousQuad: { x: number; y: number } | undefined;
  const point = (x: number, y: number, relative: boolean): { x: string; y: string } => {
    const left = relative ? currentX + x : x;
    const top = relative ? currentY + y : y;
    return localPoint(left, top);
  };
  while (index < tokens.length) {
    const token = tokens[index]!;
    if (/^[MLHVCSQTAZmlhvcsqtaz]$/u.test(token)) {
      if (token === "Z" || token === "z") {
        commands.push({ command: "close" });
        currentX = startX;
        currentY = startY;
        command = undefined;
      } else command = token;
      index += 1;
      continue;
    }
    if (command === undefined)
      throw primitiveGeometryError(undefined, "SVG path command is missing", "svg:d");
    const relative = command === command.toLowerCase();
    const argumentCounts: Record<string, number> = {
      M: 2,
      L: 2,
      H: 1,
      V: 1,
      C: 6,
      S: 4,
      Q: 4,
      T: 2,
      A: 7,
    };
    const upper = command.toUpperCase();
    const count = argumentCounts[upper] ?? 0;
    if (count === 0)
      throw primitiveGeometryError(undefined, "unsupported SVG path command", "svg:d");
    const values: number[] = [];
    for (let position = 0; position < count; position += 1) {
      const value = Number(tokens[index]);
      if (!Number.isFinite(value))
        throw primitiveGeometryError(undefined, "invalid SVG path number", "svg:d");
      values.push(value);
      index += 1;
    }
    if (upper === "M") {
      const x = values[0]!;
      const y = values[1]!;
      commands.push({ command: "moveTo", point: point(x, y, relative) });
      currentX = relative ? currentX + x : x;
      currentY = relative ? currentY + y : y;
      startX = currentX;
      startY = currentY;
      command = relative ? "l" : "L";
      previousCubic = undefined;
      previousQuad = undefined;
      continue;
    }
    if (upper === "L" || upper === "H" || upper === "V") {
      const x = upper === "V" ? currentX : relative ? currentX + values[0]! : values[0]!;
      const y =
        upper === "H"
          ? currentY
          : relative
            ? currentY + (upper === "V" ? values[0]! : values[1]!)
            : upper === "V"
              ? values[0]!
              : values[1]!;
      commands.push({ command: "lineTo", point: localPoint(x, y) });
      currentX = x;
      currentY = y;
      previousCubic = undefined;
      previousQuad = undefined;
      continue;
    }
    if (upper === "C" || upper === "S") {
      const first =
        upper === "C"
          ? {
              x: relative ? currentX + values[0]! : values[0]!,
              y: relative ? currentY + values[1]! : values[1]!,
            }
          : previousCubic
            ? { x: 2 * currentX - previousCubic.x, y: 2 * currentY - previousCubic.y }
            : { x: currentX, y: currentY };
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
          localPoint(first.x, first.y),
          localPoint(second.x, second.y),
          localPoint(end.x, end.y),
        ],
      });
      previousCubic = second;
      previousQuad = undefined;
      currentX = end.x;
      currentY = end.y;
      continue;
    }
    if (upper === "Q" || upper === "T") {
      const control =
        upper === "Q"
          ? {
              x: relative ? currentX + values[0]! : values[0]!,
              y: relative ? currentY + values[1]! : values[1]!,
            }
          : previousQuad
            ? { x: 2 * currentX - previousQuad.x, y: 2 * currentY - previousQuad.y }
            : { x: currentX, y: currentY };
      const end = {
        x: relative ? currentX + values[upper === "Q" ? 2 : 0]! : values[upper === "Q" ? 2 : 0]!,
        y: relative ? currentY + values[upper === "Q" ? 3 : 1]! : values[upper === "Q" ? 3 : 1]!,
      };
      commands.push({
        command: "quadBezTo",
        points: [localPoint(control.x, control.y), localPoint(end.x, end.y)],
      });
      previousQuad = control;
      previousCubic = undefined;
      currentX = end.x;
      currentY = end.y;
      continue;
    }
    const [radiusX, radiusY, rotation, large, sweep, x, y] = values;
    const endX = relative ? currentX + x! : x!;
    const endY = relative ? currentY + y! : y!;
    commands.push(
      ...arcToCubic(
        currentX,
        currentY,
        radiusX!,
        radiusY!,
        rotation!,
        large! !== 0,
        sweep! !== 0,
        endX,
        endY,
        localPoint,
      ),
    );
    previousCubic = undefined;
    previousQuad = undefined;
    currentX = endX;
    currentY = endY;
  }
  if (commands.length === 0) throw primitiveGeometryError(undefined, "SVG path is empty", "svg:d");
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
  localPoint: (x: number, y: number) => { x: string; y: string },
): PathCommand[] {
  if (startX === endX && startY === endY) return [];
  let rx = Math.abs(radiusX);
  let ry = Math.abs(radiusY);
  if (rx === 0 || ry === 0) return [{ command: "lineTo", point: localPoint(endX, endY) }];
  const phi = (rotationDegrees * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const left = (cos * (startX - endX)) / 2 + (sin * (startY - endY)) / 2;
  const top = (-sin * (startX - endX)) / 2 + (cos * (startY - endY)) / 2;
  const scale = (left * left) / (rx * rx) + (top * top) / (ry * ry);
  if (scale > 1) {
    const factor = Math.sqrt(scale);
    rx *= factor;
    ry *= factor;
  }
  const denominator = rx * rx * top * top + ry * ry * left * left;
  if (denominator === 0) return [{ command: "lineTo", point: localPoint(endX, endY) }];
  const factor = Math.sqrt(Math.max(0, (rx * rx * ry * ry - denominator) / denominator));
  const sign = largeArc === sweep ? -1 : 1;
  const centerLeft = (sign * factor * rx * top) / ry;
  const centerTop = (-sign * factor * ry * left) / rx;
  const centerX = cos * centerLeft - sin * centerTop + (startX + endX) / 2;
  const centerY = sin * centerLeft + cos * centerTop + (startY + endY) / 2;
  const angle = (x: number, y: number): number =>
    Math.atan2((y - centerY) / ry, (x - centerX) / rx);
  const theta = angle(startX, startY);
  let delta = angle(endX, endY) - theta;
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.ceil(Math.abs(delta) / (Math.PI / 2));
  const commands: PathCommand[] = [];
  for (let segment = 1; segment <= segments; segment += 1) {
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
        localPoint(first.x, first.y),
        localPoint(second.x, second.y),
        localPoint(target.x, target.y),
      ],
    });
  }
  return commands;
}

function parsePointList(
  points: string,
  close: boolean,
  localPoint: (x: number, y: number) => { x: string; y: string },
): PathCommand[] {
  const values = points
    .split(/[ ,;]+/u)
    .filter(Boolean)
    .map(Number);
  if (
    values.length < 4 ||
    values.length % 2 !== 0 ||
    values.some((value) => !Number.isFinite(value))
  )
    throw primitiveGeometryError(undefined, "invalid ODF point list", "draw:points");
  const commands: PathCommand[] = [];
  for (let index = 0; index < values.length; index += 2)
    commands.push({
      command: index === 0 ? "moveTo" : "lineTo",
      point: localPoint(values[index]!, values[index + 1]!),
    });
  return close ? [...commands, { command: "close" }] : commands;
}

function primitiveGeometryError(
  element: Element | undefined,
  reason: string,
  elementName: string,
): OdfSchemaError {
  return new OdfSchemaError(
    `ODF primitive geometry: ${reason}`,
    "content.xml",
    element?.name ? `/draw:page/${element.name}` : "/draw:page",
    elementName,
    reason,
  );
}

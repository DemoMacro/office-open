import type { TextBodyOptions } from "@office-open/core";
import type {
  GroupOptions,
  PictureOptions,
  ShapeOptions,
  SlideChild,
  TableOptions,
} from "@office-open/pptx";

import {
  LegacyPowerPointError,
  RecordType,
  findDescendant,
  findDirect,
  readInt16,
  readInt32,
  type RecordNode,
} from "../records";
import { TABLE_BOUND_TOLERANCE } from "./constants";
import type { DrawingContext, LegacyPicture, TableCell } from "./models";
import {
  applyClientHyperlinks,
  collectEmbeddedText,
  createTextBody,
  masterUnitsToEmu,
} from "./text";
export function readDrawingChildren(
  view: DataView,
  drawing: RecordNode,
  context: DrawingContext,
): SlideChild[] {
  const pageContainer = findDirect(drawing, RecordType.escherDrawingContainer);
  const shapeGroups = pageContainer
    ? pageContainer.children.filter((child) => child.type === RecordType.escherShapeGroupContainer)
    : [];
  const children: SlideChild[] = [];
  for (const shapeGroup of shapeGroups) {
    children.push(...parseShapeGroup(view, shapeGroup, context, 0));
  }
  return children;
}

function parseShapeGroup(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
  depth: number,
): SlideChild[] {
  if (depth > 16) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint drawing: group depth limit exceeded",
    );
  }

  const children: SlideChild[] = [];
  const childRecords = container.children;
  const metadata = childRecords.find(
    (child) =>
      child.type === RecordType.escherShapeContainer &&
      findDirect(child, RecordType.escherShapeGroup) !== undefined,
  );

  for (const child of childRecords) {
    if (child.type === RecordType.escherShapeGroupContainer) {
      children.push(...parseShapeGroup(view, child, context, depth + 1));
      continue;
    }
    if (child.type !== RecordType.escherShapeContainer || child === metadata) continue;
    const shape = parseShape(view, child, context);
    if (!shape) continue;
    if ("data" in shape) children.push({ picture: shape });
    else children.push({ shape });
  }

  if (children.length === 0) return [];
  if (depth === 0 && children.length === 1) {
    const only = children[0]!;
    if ("table" in only) return children;
    if ("group" in only) {
      const nested = only.group.children;
      if (nested.length === 1 && "table" in nested[0]!) return nested;
    }
  }

  const table = detectGroupedTable(view, childRecords, context);
  if (table) return [{ table }];

  const group = parseGroupShape(view, metadata, children);
  return group ? [{ group }] : children;
}

function detectGroupedTable(
  view: DataView,
  records: readonly RecordNode[],
  context: DrawingContext,
): TableOptions | undefined {
  const metadata = records.find(
    (record) =>
      record.type === RecordType.escherShapeContainer &&
      findDirect(record, RecordType.escherShapeGroup) !== undefined,
  );
  const frame = readTableFrame(view, metadata);
  const groupBounds = readGroupBounds(view, metadata);
  const shapeRecords = records.filter(
    (record) => record.type === RecordType.escherShapeContainer && record !== metadata,
  );
  if (shapeRecords.some((record) => !findDirect(record, RecordType.escherChildAnchor))) {
    return undefined;
  }

  const bounds = shapeRecords.flatMap(
    (
      record,
    ): {
      shape: ShapeOptions;
      left: number;
      top: number;
      right: number;
      bottom: number;
    }[] => {
      const shape = parseShape(view, record, context);
      const anchor = readAnchorBounds(view, record);
      if (!shape || !("textBody" in shape) || !anchor) return [];
      return [{ shape, ...anchor }];
    },
  );
  if (!frame || frame.width <= 0 || frame.height <= 0) return undefined;
  if (bounds.some((item) => item.left >= item.right || item.top >= item.bottom)) return undefined;
  if (groupBounds) {
    const childLeft = Math.min(groupBounds.childLeft, groupBounds.childRight);
    const childTop = Math.min(groupBounds.childTop, groupBounds.childBottom);
    const childRight = Math.max(groupBounds.childLeft, groupBounds.childRight);
    const childBottom = Math.max(groupBounds.childTop, groupBounds.childBottom);
    if (
      bounds.some(
        (item) =>
          item.left < childLeft - TABLE_BOUND_TOLERANCE ||
          item.top < childTop - TABLE_BOUND_TOLERANCE ||
          item.right > childRight + TABLE_BOUND_TOLERANCE ||
          item.bottom > childBottom + TABLE_BOUND_TOLERANCE,
      )
    ) {
      return undefined;
    }
  }

  const lefts = uniqueSorted(bounds.map((item) => item.left));
  const tops = uniqueSorted(bounds.map((item) => item.top));
  const rightEdges = uniqueSorted(bounds.map((item) => item.right));
  const bottomEdges = uniqueSorted(bounds.map((item) => item.bottom));
  const columnEdges = uniqueSorted([...lefts, ...rightEdges]);
  const rowEdges = uniqueSorted([...tops, ...bottomEdges]);
  if (
    columnEdges.length < 3 ||
    tops.length < 2 ||
    columnEdges.length !== lefts.length + 1 ||
    rowEdges.length !== tops.length + 1 ||
    lefts.some((left, index) => !near(left, columnEdges[index]!)) ||
    rightEdges.some((right, index) => !near(right, columnEdges[index + 1]!)) ||
    tops.some((top, index) => !near(top, rowEdges[index]!)) ||
    bottomEdges.some((bottom, index) => !near(bottom, rowEdges[index + 1]!))
  ) {
    return undefined;
  }

  const columnCount = columnEdges.length - 1;
  const rowCount = rowEdges.length - 1;
  const occupied = Array.from({ length: rowCount }, () => Array<boolean>(columnCount).fill(false));
  const cells: (TableCell | undefined)[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: columnCount }, () => undefined),
  );

  for (const item of bounds.sort(
    (first, second) => first.top - second.top || first.left - second.left,
  )) {
    const columnStart = columnEdges.findIndex((edge) => near(edge, item.left));
    const columnEnd = columnEdges.findIndex((edge) => near(edge, item.right));
    const rowStart = rowEdges.findIndex((edge) => near(edge, item.top));
    const rowEnd = rowEdges.findIndex((edge) => near(edge, item.bottom));
    const columnSpan = columnEnd - columnStart;
    const rowSpan = rowEnd - rowStart;
    if (columnStart < 0 || rowStart < 0 || columnSpan <= 0 || rowSpan <= 0) return undefined;

    let firstColumn = -1;
    let firstRow = -1;
    outer: for (let row = rowStart; row < rowEnd; row += 1) {
      for (let column = columnStart; column < columnEnd; column += 1) {
        if (!occupied[row]![column]) {
          firstRow = row;
          firstColumn = column;
          break outer;
        }
      }
    }
    if (firstRow < 0 || firstColumn < 0) return undefined;
    for (let row = firstRow; row < firstRow + rowSpan; row += 1) {
      for (let column = firstColumn; column < firstColumn + columnSpan; column += 1) {
        if (occupied[row]![column]) return undefined;
        occupied[row]![column] = true;
      }
    }
    cells[firstRow]![firstColumn] = {
      ...tableCell(item.shape),
      ...(columnSpan > 1 ? { columnSpan } : {}),
      ...(rowSpan > 1 ? { rowSpan } : {}),
    };
  }
  if (occupied.some((row) => row.some((cell) => !cell))) return undefined;

  const columnWidths = columnEdges
    .slice(0, -1)
    .map((edge, index) => masterUnitsToEmu(columnEdges[index + 1]! - edge));
  const rowHeights = rowEdges
    .slice(0, -1)
    .map((edge, index) => masterUnitsToEmu(rowEdges[index + 1]! - edge));

  const rows: NonNullable<TableOptions["rows"]> = [];
  for (const [index, height] of rowHeights.entries()) {
    if (height <= 0) return undefined;
    rows.push({ height, cells: cells[index]!.filter((cell) => cell !== undefined) });
  }

  return { ...frame, rows, columnWidths };
}

function readAnchorBounds(
  view: DataView,
  container: RecordNode,
): { left: number; top: number; right: number; bottom: number } | undefined {
  const clientAnchor = findDirect(container, RecordType.escherClientAnchor);
  if (clientAnchor) {
    if (clientAnchor.length < 8) return undefined;
    return {
      top: readInt16(view, clientAnchor, 0),
      left: readInt16(view, clientAnchor, 2),
      right: readInt16(view, clientAnchor, 4),
      bottom: readInt16(view, clientAnchor, 6),
    };
  }

  const childAnchor = findDirect(container, RecordType.escherChildAnchor);
  if (!childAnchor || childAnchor.length < 16) return undefined;
  return {
    left: readInt32(view, childAnchor, 0),
    top: readInt32(view, childAnchor, 4),
    right: readInt32(view, childAnchor, 8),
    bottom: readInt32(view, childAnchor, 12),
  };
}

function uniqueSorted(values: readonly number[]): number[] {
  return [...new Set(values)].sort((first, second) => first - second);
}

function near(first: number, second: number): boolean {
  return Math.abs(first - second) <= TABLE_BOUND_TOLERANCE;
}

export function textBodyToText(body: TextBodyOptions | undefined): string {
  if (body?.text !== undefined) return body.text;
  return (body?.paragraphs ?? [])
    .map((paragraph) =>
      typeof paragraph === "string"
        ? paragraph
        : (paragraph.children ?? [])
            .map((child) =>
              typeof child === "string" ? child : "text" in child ? (child.text ?? "") : "",
            )
            .join(""),
    )
    .join("\n");
}

function parseGroupShape(
  view: DataView,
  metadata: RecordNode | undefined,
  children: SlideChild[],
): GroupOptions | undefined {
  const transform = readGroupTransform(view, metadata);
  if (!transform) return undefined;
  return { ...transform, children };
}

interface GroupBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly childLeft: number;
  readonly childTop: number;
  readonly childRight: number;
  readonly childBottom: number;
}

interface TableFrame {
  readonly id?: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

function readGroupTransform(
  view: DataView,
  metadata: RecordNode | undefined,
):
  | Pick<
      GroupOptions,
      | "id"
      | "x"
      | "y"
      | "width"
      | "height"
      | "childOffsetX"
      | "childOffsetY"
      | "childExtentWidth"
      | "childExtentHeight"
    >
  | undefined {
  const bounds = readGroupBounds(view, metadata);
  if (!bounds) return undefined;
  const shapeRecord = metadata && findDirect(metadata, RecordType.escherShape);
  return {
    ...(shapeRecord && shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
    childOffsetX: masterUnitsToEmu(Math.min(bounds.childLeft, bounds.childRight)),
    childOffsetY: masterUnitsToEmu(Math.min(bounds.childTop, bounds.childBottom)),
    childExtentWidth: masterUnitsToEmu(Math.abs(bounds.childRight - bounds.childLeft)),
    childExtentHeight: masterUnitsToEmu(Math.abs(bounds.childBottom - bounds.childTop)),
  };
}

function readGroupBounds(
  view: DataView,
  metadata: RecordNode | undefined,
): GroupBounds | undefined {
  const groupAtom = metadata && findDirect(metadata, RecordType.escherShapeGroup);
  if (!groupAtom || groupAtom.length < 16) return undefined;
  const bounds = {
    left: readInt32(view, groupAtom, 0),
    top: readInt32(view, groupAtom, 4),
    right: readInt32(view, groupAtom, 8),
    bottom: readInt32(view, groupAtom, 12),
  };
  if (groupAtom.length < 32) {
    return {
      ...bounds,
      childLeft: bounds.left,
      childTop: bounds.top,
      childRight: bounds.right,
      childBottom: bounds.bottom,
    };
  }
  return {
    ...bounds,
    childLeft: readInt32(view, groupAtom, 16),
    childTop: readInt32(view, groupAtom, 20),
    childRight: readInt32(view, groupAtom, 24),
    childBottom: readInt32(view, groupAtom, 28),
  };
}

function readTableFrame(view: DataView, metadata: RecordNode | undefined): TableFrame | undefined {
  const bounds = readGroupBounds(view, metadata);
  if (!bounds) return undefined;
  const shapeRecord = metadata && findDirect(metadata, RecordType.escherShape);
  return {
    ...(shapeRecord && shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };
}

function tableCell(shape: ShapeOptions): TableCell {
  const paragraphs = shape.textBody?.paragraphs;
  if (paragraphs !== undefined && paragraphs.length > 1) return { children: paragraphs };
  return { text: textBodyToText(shape.textBody) };
}

function parseShape(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): ShapeOptions | PictureOptions | undefined {
  const shapeRecord = findDirect(container, RecordType.escherShape);
  if (!shapeRecord) return undefined;
  const bounds = readAnchorBounds(view, container);
  if (!bounds) return undefined;
  const picture =
    shapeRecord.instance === 75 ? readShapePicture(view, container, context.pictures) : undefined;
  const text = readContainerText(view, container, context);
  if (!text && !picture) return undefined;

  const position = {
    ...(shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };
  if (picture && !text) return { ...position, type: picture.type, data: picture.data };
  if (picture) {
    return {
      ...position,
      textBox: true,
      properties: {
        fill: { type: "blip", data: picture.data, imageType: picture.type },
      },
      textBody: text,
    };
  }
  return {
    ...position,
    textBox: true,
    textBody: text,
  };
}

function readShapePicture(
  view: DataView,
  container: RecordNode,
  pictures: readonly (LegacyPicture | undefined)[],
): LegacyPicture | undefined {
  const options = findDirect(container, RecordType.escherShapeProperties);
  if (!options) return undefined;
  const reference = findShapePictureReference(view, options);
  if (reference === undefined || reference <= 0 || reference >= pictures.length) return undefined;
  return pictures[reference];
}

function findShapePictureReference(view: DataView, options: RecordNode): number | undefined {
  const headerEnd = options.offset + 8;
  const propertyCount = options.instance;
  const propertiesEnd = headerEnd + propertyCount * 6;
  if (propertiesEnd > options.end) return undefined;
  let cursor = options.offset + 8;
  let result: number | undefined;
  for (let index = 0; index < propertyCount; index += 1) {
    const packedId = view.getUint16(cursor, true);
    const value = view.getUint32(cursor + 2, true);
    const propertyId = packedId & 0x3fff;
    const isBlip = (packedId & 0x4000) !== 0;
    if (propertyId === 0x0104 && isBlip) result = value;
    cursor += 6;
  }
  return result;
}

export function readContainerText(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): TextBodyOptions | undefined {
  const clientTextbox = findDirect(container, RecordType.escherClientTextbox);
  if (!clientTextbox) return undefined;

  const embeddedText = collectEmbeddedText(view, clientTextbox);
  if (embeddedText.length > 0) {
    const body = createTextBody(embeddedText);
    return applyClientHyperlinks(body, view, clientTextbox, context.hyperlinks);
  }

  const reference = findDescendant(clientTextbox, RecordType.outlineTextReference);
  if (reference && reference.length >= 4) {
    const index = readInt32(view, reference, 0);
    const entry = context.entries[index];
    if (entry) {
      context.usedEntryIndexes.add(index);
      return createTextBody([entry.text]);
    }
  }
  return undefined;
}

import type { TextBodyOptions } from "@office-open/core";
import type { ShapePropertiesOptions } from "@office-open/core/drawing";
import type { GroupOptions, ShapeOptions, SlideChild, TableOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";
import {
  RecordType,
  findDescendant,
  findDirect,
  readInt16,
  readInt32,
  type RecordNode,
} from "../records";
import { TABLE_BOUND_TOLERANCE } from "./constants";
import {
  CONNECTOR_TYPE_PRESETS,
  SHAPE_TYPE_PRESETS,
  SHAPE_TYPE_WARPS,
  readShapeFlags,
  readShapeProperties,
  type EscherShapeStyle,
} from "./escher-properties";
import type { DrawingContext, TableCell, TextEntry } from "./models";
import { readShapePlaceholder } from "./placeholders";
import {
  applyClientHyperlinks,
  collectEmbeddedTextEntries,
  legacyPlaceholder,
  createTextBody,
  masterUnitsToEmu,
} from "./text";

type EndpointShapeProperties = Pick<ShapePropertiesOptions, "fill" | "outline" | "effects">;
export function readDrawingChildren(
  view: DataView,
  drawing: RecordNode,
  context: DrawingContext,
): SlideChild[] {
  const pageContainer = findDirect(drawing, RecordType.escherDrawingContainer);
  const childContainers = pageContainer
    ? pageContainer.children.filter(
        (child) =>
          child.type === RecordType.escherShapeGroupContainer ||
          child.type === RecordType.escherShapeContainer,
      )
    : [];
  const children: SlideChild[] = [];
  for (const childContainer of childContainers) {
    if (childContainer.type === RecordType.escherShapeContainer) {
      const shape = parseShape(view, childContainer, context);
      if (shape) children.push(shape);
      continue;
    }
    children.push(...parseShapeGroup(view, childContainer, context, 0));
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
    throw new PptParseError("Corrupt legacy PowerPoint drawing: group depth limit exceeded");
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
    if (shape) children.push(shape);
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
      const child = parseShape(view, record, context);
      const anchor = readAnchorBounds(view, record);
      if (!child || !("shape" in child) || !child.shape.textBody || !anchor) return [];
      return [{ shape: child.shape, ...anchor }];
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
): SlideChild | undefined {
  const shapeRecord = findDirect(container, RecordType.escherShape);
  if (!shapeRecord) return undefined;
  const bounds = readAnchorBounds(view, container);
  if (!bounds) return undefined;
  const optionsRecord = findDirect(container, RecordType.escherShapeProperties);
  const style = optionsRecord ? readShapeProperties(view, optionsRecord) : {};
  const flags = readShapeFlags(view, shapeRecord);
  const { body: text, entry: textEntry } = readContainerText(view, container, context);
  const placeholderMetadata = readShapePlaceholder(view, container);
  const placeholder =
    placeholderMetadata === null
      ? undefined
      : (placeholderMetadata?.placeholder ?? legacyPlaceholder(textEntry?.placeholderType));
  const position = {
    ...(shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };

  if (shapeRecord.instance === 75) {
    return parsePictureShape(position, style, flags, text, placeholderMetadata, context);
  }
  if (shapeRecord.instance === 20) {
    const lineProperties = endpointShapeProperties(style, context);
    return {
      line: {
        ...(position.id !== undefined ? { id: position.id } : {}),
        x1: masterUnitsToEmu(bounds.left),
        y1: masterUnitsToEmu(bounds.top),
        x2: masterUnitsToEmu(bounds.right),
        y2: masterUnitsToEmu(bounds.bottom),
        ...(style.rotation !== undefined ? { rotation: style.rotation } : {}),
        ...(Object.keys(lineProperties).length > 0 ? { properties: lineProperties } : {}),
        ...(text ? { textBody: text } : {}),
      },
    };
  }
  if (CONNECTOR_TYPE_PRESETS[shapeRecord.instance] !== undefined || flags.connector) {
    const preset = CONNECTOR_TYPE_PRESETS[shapeRecord.instance];
    const properties = endpointShapeProperties(style, context);
    return {
      connector: {
        ...(position.id !== undefined ? { id: position.id } : {}),
        x1: masterUnitsToEmu(bounds.left),
        y1: masterUnitsToEmu(bounds.top),
        x2: masterUnitsToEmu(bounds.right),
        y2: masterUnitsToEmu(bounds.bottom),
        ...(style.rotation !== undefined ? { rotation: style.rotation } : {}),
        ...(preset !== undefined || Object.keys(properties).length > 0
          ? { properties: { ...(preset !== undefined ? { geometry: preset } : {}), ...properties } }
          : {}),
      },
    };
  }

  const properties: ShapeOptions["properties"] = {};
  const geometry = SHAPE_TYPE_PRESETS[shapeRecord.instance];
  if (geometry !== undefined) properties.geometry = geometry;
  const textWarp = SHAPE_TYPE_WARPS[shapeRecord.instance];
  const wordArtText =
    text && textWarp
      ? {
          ...text,
          bodyProperties: {
            ...text.bodyProperties,
            fromWordArt: true,
            prstTxWarp: { preset: textWarp },
          },
        }
      : text;
  if (shapeRecord.instance === 0) {
    if (!style.customGeometry) {
      throw new PptParseError("Legacy PowerPoint freeform has no canonical geometry", {
        part: "record",
        recordType: shapeRecord.type,
        recordName: "OfficeArtFSP",
        byteRange: [shapeRecord.offset, shapeRecord.end],
        reason: "missing-freeform-geometry",
      });
    }
    properties.customGeometry = style.customGeometry;
  } else if (
    geometry === undefined &&
    shapeRecord.instance !== 202 &&
    SHAPE_TYPE_WARPS[shapeRecord.instance] === undefined
  ) {
    throw new PptParseError(`Unsupported legacy PowerPoint shape type ${shapeRecord.instance}`, {
      part: "record",
      recordType: shapeRecord.type,
      recordName: "OfficeArtFSP",
      byteRange: [shapeRecord.offset, shapeRecord.end],
      reason: "unsupported-shape-type",
    });
  }
  const fill = resolveFill(style, context);
  if (fill) properties.fill = fill;
  if (style.outline) properties.outline = style.outline;
  if (style.effects) properties.effects = style.effects;
  return {
    shape: {
      ...position,
      ...(style.rotation !== undefined ? { rotation: style.rotation } : {}),
      ...(flags.flipHorizontal ? { flipHorizontal: true } : {}),
      ...(flags.flipVertical ? { flipVertical: true } : {}),
      ...(Object.keys(properties).length > 0 ? { properties } : {}),
      ...(wordArtText ? { textBox: true, textBody: wordArtText } : {}),
      ...(placeholder !== undefined ? { placeholder } : {}),
      ...(placeholderMetadata?.placeholderSize !== undefined
        ? { placeholderSize: placeholderMetadata.placeholderSize }
        : textEntry?.placeholderSize !== undefined
          ? { placeholderSize: textEntry.placeholderSize }
          : {}),
      ...(placeholderMetadata?.placeholderIndex !== undefined
        ? { placeholderIndex: placeholderMetadata.placeholderIndex }
        : {}),
    },
  };
}

function parsePictureShape(
  position: { id?: number; x: number; y: number; width: number; height: number },
  style: EscherShapeStyle,
  flags: ReturnType<typeof readShapeFlags>,
  text: TextBodyOptions | undefined,
  placeholderMetadata: ReturnType<typeof readShapePlaceholder>,
  context: DrawingContext,
): SlideChild | undefined {
  const reference = style.blipReference;
  const picture =
    reference !== undefined && reference > 0 && reference < context.pictures.length
      ? context.pictures[reference]
      : undefined;
  if (!picture) {
    return text
      ? {
          shape: {
            ...position,
            textBox: true,
            textBody: text,
            ...placeholderOptions(placeholderMetadata),
          },
        }
      : undefined;
  }
  if (!text) {
    return {
      picture: {
        ...position,
        ...(flags.flipHorizontal ? { flipHorizontal: true } : {}),
        ...(flags.flipVertical ? { flipVertical: true } : {}),
        ...(style.rotation !== undefined ? { rotation: style.rotation } : {}),
        type: picture.type,
        data: picture.data,
        ...placeholderOptions(placeholderMetadata),
      },
    };
  }
  const fill =
    resolveFill(style, context) ??
    ({
      type: "blip",
      data: picture.data,
      imageType: picture.type,
    } as const);
  return {
    shape: {
      ...position,
      textBox: true,
      properties: { fill },
      textBody: text,
    },
  };
}

function endpointShapeProperties(
  style: EscherShapeStyle,
  context: DrawingContext,
): EndpointShapeProperties {
  const properties: EndpointShapeProperties = {};
  const fill = resolveFill(style, context);
  if (fill) properties.fill = fill;
  if (style.outline) properties.outline = style.outline;
  if (style.effects) properties.effects = style.effects;
  return properties;
}

function placeholderOptions(
  metadata: ReturnType<typeof readShapePlaceholder>,
): ShapeOptions | Record<string, never> {
  if (!metadata) return {};
  return {
    ...(metadata.placeholder !== undefined ? { placeholder: metadata.placeholder } : {}),
    ...(metadata.placeholderSize !== undefined
      ? { placeholderSize: metadata.placeholderSize }
      : {}),
    ...(metadata.placeholderIndex !== undefined
      ? { placeholderIndex: metadata.placeholderIndex }
      : {}),
  };
}

function resolveFill(style: EscherShapeStyle, context: DrawingContext) {
  if (style.fillType === 5) {
    const reference = style.fillBlipReference;
    const picture =
      reference !== undefined && reference > 0 && reference < context.pictures.length
        ? context.pictures[reference]
        : undefined;
    return picture
      ? ({ type: "blip", data: picture.data, imageType: picture.type } as const)
      : undefined;
  }
  return style.fill;
}

interface ContainerText {
  readonly body: TextBodyOptions | undefined;
  readonly entry: TextEntry | undefined;
}

export function readContainerText(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): ContainerText {
  const clientTextbox = findDirect(container, RecordType.escherClientTextbox);
  if (!clientTextbox) return { body: undefined, entry: undefined };

  const embeddedText = collectEmbeddedTextEntries(view, clientTextbox, context.fonts);
  if (embeddedText.length > 0) {
    const body = createTextBody(embeddedText);
    return {
      body: applyClientHyperlinks(body, view, clientTextbox, context.hyperlinks),
      entry: embeddedText[0],
    };
  }

  const reference = findDescendant(clientTextbox, RecordType.outlineTextReference);
  if (reference && reference.length >= 4) {
    const index = readInt32(view, reference, 0);
    const entry = context.entries[index];
    if (entry) {
      context.usedEntryIndexes.add(index);
      return { body: createTextBody([entry]), entry };
    }
  }
  return { body: undefined, entry: undefined };
}

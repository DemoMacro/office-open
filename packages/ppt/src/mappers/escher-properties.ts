import type {
  CustomGeometryOptions,
  EffectListOptions,
  FillOptions,
  GradientStopOptions,
  LineEndOptions,
  OutlineOptions,
  RgbColorOptions,
  PathCommand,
  PathOptions,
  ShapeType,
  TextShapePreset,
} from "@office-open/core/drawing";

import { PptParseError } from "../errors";
import type { RecordNode } from "../records";

/** Escher shape-instance → OOXML preset geometry for ordinary autoshapes. */
export const SHAPE_TYPE_PRESETS: Readonly<Record<number, ShapeType>> = {
  1: "rect",
  2: "roundRect",
  3: "ellipse",
  4: "diamond",
  5: "triangle",
  6: "rtTriangle",
  7: "parallelogram",
  8: "trapezoid",
  9: "hexagon",
  10: "octagon",
  11: "plus",
  12: "star5",
  13: "rightArrow",
  15: "homePlate",
  19: "arc",
  21: "plaque",
  22: "can",
  23: "donut",
  193: "actionButtonForwardNext",
};

interface GeometryBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

interface ComplexPointReader {
  readonly view: DataView;
  offset: number;
  readonly pointSize: number;
}

const MSO_PATH_LINE_TO = 0;
const MSO_PATH_CURVE_TO = 1;
const MSO_PATH_MOVE_TO = 2;
const MSO_PATH_CLOSE = 3;
const MSO_PATH_END = 4;

export const SHAPE_TYPE_WARPS: Readonly<Record<number, TextShapePreset>> = {
  24: "textPlain",
  27: "textCurveUp",
  28: "textWave1",
  29: "textRingInside",
  30: "textArchUp",
  31: "textRingOutside",
  137: "textStop",
  138: "textTriangle",
  139: "textTriangleInverted",
  140: "textChevron",
  141: "textChevronInverted",
  142: "textRingInside",
  143: "textRingOutside",
  144: "textArchUp",
  145: "textArchDown",
  146: "textCircle",
  147: "textButton",
  148: "textArchUpPour",
  149: "textArchDownPour",
  150: "textCirclePour",
  151: "textButtonPour",
  152: "textCurveUp",
  153: "textCurveDown",
  154: "textCascadeUp",
  155: "textCascadeDown",
  156: "textWave1",
  157: "textWave2",
  159: "textWave4",
  160: "textInflate",
  161: "textDeflate",
  162: "textInflateBottom",
  163: "textDeflateBottom",
  164: "textInflateTop",
  165: "textDeflateTop",
  166: "textDeflateInflate",
  167: "textDeflateInflateDeflate",
  168: "textFadeRight",
  169: "textFadeLeft",
  170: "textFadeUp",
  171: "textFadeDown",
  172: "textSlantUp",
  173: "textSlantDown",
  174: "textCanUp",
  175: "textCanDown",
};

/** Escher connector shape-instance → OOXML preset connector geometry. */
export const CONNECTOR_TYPE_PRESETS: Readonly<Record<number, ShapeType>> = {
  32: "straightConnector1",
  33: "bentConnector2",
  34: "bentConnector3",
  35: "bentConnector4",
  36: "bentConnector5",
  37: "curvedConnector2",
  38: "curvedConnector3",
  39: "curvedConnector4",
  40: "curvedConnector5",
};

/** OfficeArtFSP flags: fFlipH/fFlipV/fConnector live in the body flags word. */
export interface EscherShapeFlags {
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
  readonly connector: boolean;
}

/** Canonical projection of one OfficeArtFOPT property matrix. */
export interface EscherShapeStyle {
  /** 0x0104 blip reference for the shape's own picture (instance 75). */
  blipReference?: number;
  /** 0x0186 blip reference when fill type is 5 (picture fill). */
  fillBlipReference?: number;
  /** Raw 0x0180 fill type so callers can resolve blip fills. */
  fillType?: number;
  fill?: FillOptions;
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  customGeometry?: CustomGeometryOptions;
  /** Rotation in degrees (0x0004 is 1/65536-degree fixed point). */
  rotation?: number;
}

interface CollectedProperty {
  readonly value: number;
  readonly complex?: DataView;
  readonly blip: boolean;
  readonly offset: number;
}

const LINE_DASHES: Readonly<Record<number, NonNullable<OutlineOptions["dash"]>>> = {
  0: "solid",
  1: "sysDash",
  2: "sysDot",
  3: "sysDashDot",
  4: "sysDashDotDot",
  5: "dot",
  6: "dash",
  7: "lgDashDot",
  8: "lgDashDotDot",
  9: "dashDot",
};

const ARROW_HEADS: Readonly<Record<number, NonNullable<LineEndOptions["type"]>>> = {
  1: "triangle",
  2: "stealth",
  3: "diamond",
  4: "oval",
};

const ARROW_SIZES: Readonly<
  Record<number, NonNullable<LineEndOptions["width"] | LineEndOptions["length"]>>
> = {
  0: "small",
  1: "medium",
  2: "large",
};

const KNOWN_PROPERTY_IDS = new Set([
  0x0004, 0x0104, 0x0180, 0x0181, 0x0182, 0x0183, 0x0186, 0x018b, 0x01c0, 0x01cb, 0x01ce, 0x01d1,
  0x01d2, 0x01d3, 0x01d4, 0x01d5, 0x01d6, 0x0200, 0x0201, 0x0205, 0x0206,
]);

/**
 * Property ids whose visual impact the legacy parser intentionally does not
 * project (text layout, blip crops, geometry adjustment values, compound line
 * variants, shadow detail, perspective, and 3D). Known but unsupported visual
 * types such as pattern fills throw instead.
 */
const isIgnorableProperty = (id: number): boolean =>
  id < 0x0004 ||
  (id >= 0x0005 && id < 0x0104) ||
  (id > 0x0104 && id < 0x0180) ||
  id === 0x0184 ||
  id === 0x0185 ||
  (id >= 0x0187 && id < 0x018b) ||
  (id > 0x018b && id < 0x01c0) ||
  (id >= 0x01c1 && id < 0x01cb) ||
  (id >= 0x01cc && id <= 0x01d0) ||
  (id >= 0x01d7 && id < 0x0200) ||
  (id >= 0x0202 && id <= 0x0204) ||
  (id >= 0x0207 && id < 0x0400);

export function readShapeFlags(view: DataView, shape: RecordNode): EscherShapeFlags {
  if (shape.length < 8) {
    return { flipHorizontal: false, flipVertical: false, connector: false };
  }
  const flags = view.getUint32(shape.offset + 12, true);
  return {
    flipHorizontal: (flags & 0x40) !== 0,
    flipVertical: (flags & 0x80) !== 0,
    connector: (flags & 0x100) !== 0,
  };
}

export function readShapeProperties(view: DataView, options: RecordNode): EscherShapeStyle {
  const entriesEnd = options.offset + 8 + options.instance * 6;
  if (entriesEnd > options.end) {
    throw new PptParseError("Corrupt legacy PowerPoint shape property table", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      byteRange: [options.offset, entriesEnd],
      reason: "invalid-property-table",
    });
  }
  const collected = new Map<number, CollectedProperty>();
  let complexOffset = entriesEnd;
  for (let index = 0; index < options.instance; index += 1) {
    const offset = options.offset + 8 + index * 6;
    const packedId = view.getUint16(offset, true);
    const value = view.getUint32(offset + 2, true);
    let complex: DataView | undefined;
    if ((packedId & 0x8000) !== 0) {
      const complexEnd = complexOffset + value;
      if (complexEnd > options.end) {
        throw new PptParseError("Corrupt legacy PowerPoint complex shape property", {
          part: "record",
          recordType: options.type,
          recordName: "OfficeArtFOPT",
          propertyId: packedId & 0x3fff,
          byteRange: [complexOffset, complexEnd],
          reason: "truncated-complex-property",
        });
      }
      complex = new DataView(view.buffer, view.byteOffset + complexOffset, value);
      complexOffset = complexEnd;
    }
    collected.set(packedId & 0x3fff, {
      value,
      complex,
      blip: (packedId & 0x4000) !== 0,
      offset,
    });
  }
  const style: EscherShapeStyle = {};
  readRotation(collected, style);
  readBlipReference(collected, style);
  readCustomGeometry(collected, style, options);
  readFill(collected, style, options);
  readOutline(collected, style);
  readShadow(collected, style);
  rejectUnknownProperties(collected, options);
  return style;
}

function readSignedProperty(
  collected: ReadonlyMap<number, CollectedProperty>,
  propertyId: number,
  options: RecordNode,
): number {
  const property = collected.get(propertyId);
  if (!property) {
    throw new PptParseError("Missing legacy PowerPoint custom geometry property", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId,
      reason: "missing-custom-geometry-property",
    });
  }
  if (property.complex) {
    if (property.complex.byteLength < 4) {
      throw new PptParseError("Truncated legacy PowerPoint custom geometry bounds", {
        part: "record",
        recordType: options.type,
        recordName: "OfficeArtFOPT",
        propertyId,
        length: property.complex.byteLength,
        byteRange: [property.offset, property.offset + 6],
        reason: "truncated-geometry-bounds",
      });
    }
    return property.complex.getInt32(0, true);
  }
  return property.value | 0;
}

function readCustomGeometry(
  collected: ReadonlyMap<number, CollectedProperty>,
  style: EscherShapeStyle,
  options: RecordNode,
): void {
  const vertices = collected.get(325);
  const segments = collected.get(326);
  if (!vertices && !segments) return;
  if (!vertices?.complex || !segments?.complex) {
    throw new PptParseError("Missing legacy PowerPoint freeform geometry data", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: vertices?.complex ? 326 : 325,
      byteRange: [options.offset, options.end],
      reason: "missing-freeform-geometry-data",
    });
  }
  if (vertices.complex.byteLength < 6) {
    throw new PptParseError("Truncated legacy PowerPoint freeform vertices", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 325,
      length: vertices.complex.byteLength,
      reason: "truncated-freeform-vertices",
    });
  }
  if (segments.complex.byteLength < 4) {
    throw new PptParseError("Truncated legacy PowerPoint freeform segments", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 326,
      length: segments.complex.byteLength,
      reason: "truncated-freeform-segments",
    });
  }

  const bounds: GeometryBounds = {
    left: readSignedProperty(collected, 320, options),
    top: readSignedProperty(collected, 321, options),
    right: readSignedProperty(collected, 322, options),
    bottom: readSignedProperty(collected, 323, options),
  };
  const width = bounds.right - bounds.left;
  const height = bounds.bottom - bounds.top;
  if (width <= 0 || height <= 0) {
    throw new PptParseError("Invalid legacy PowerPoint freeform geometry bounds", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      byteRange: [options.offset, options.end],
      reason: "invalid-freeform-bounds",
    });
  }

  const pointSize = vertices.complex.getUint16(4, true);
  const reader: ComplexPointReader = {
    view: vertices.complex,
    offset: 6,
    pointSize,
  };
  const segmentCount = segments.complex.getUint16(0, true);
  let segmentOffset = 4;
  const pathList: PathOptions[] = [];
  let commands: PathCommand[] = [];

  for (let index = 0; index < segmentCount; index += 1) {
    if (segmentOffset + 4 > segments.complex.byteLength) {
      throw new PptParseError("Truncated legacy PowerPoint freeform segment", {
        part: "record",
        recordType: options.type,
        recordName: "OfficeArtFOPT",
        propertyId: 326,
        byteRange: [segmentOffset, segmentOffset + 4],
        reason: "truncated-freeform-segment",
      });
    }
    const segment = segments.complex.getUint16(segmentOffset, true);
    const type = (segment & 0xe000) >>> 13;
    const count = segment & 0x03ff;
    segmentOffset += 4;

    if (type === MSO_PATH_MOVE_TO) {
      commands.push({ command: "moveTo", point: readPoint(reader, options) });
      continue;
    }
    if (type === MSO_PATH_LINE_TO) {
      for (let point = 0; point < count; point += 1) {
        commands.push({ command: "lineTo", point: readPoint(reader, options) });
      }
      continue;
    }
    if (type === MSO_PATH_CURVE_TO) {
      for (let curve = 0; curve < count; curve += 1) {
        commands.push({
          command: "cubicBezTo",
          points: [
            readPoint(reader, options),
            readPoint(reader, options),
            readPoint(reader, options),
          ],
        });
      }
      continue;
    }
    if (type === MSO_PATH_CLOSE) {
      commands.push({ command: "close" });
      continue;
    }
    if (type === MSO_PATH_END) {
      if (commands.length > 0) pathList.push({ w: width, h: height, commands });
      commands = [];
      continue;
    }
    throw new PptParseError(`Unsupported legacy PowerPoint path segment type ${type}`, {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 326,
      reason: "unsupported-freeform-segment-type",
    });
  }
  if (commands.length > 0) pathList.push({ w: width, h: height, commands });
  if (pathList.length === 0) {
    throw new PptParseError("Legacy PowerPoint freeform has no geometry paths", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 326,
      reason: "empty-freeform-geometry",
    });
  }
  style.customGeometry = { pathList };
}

function readPoint(reader: ComplexPointReader, options: RecordNode): { x: string; y: string } {
  const componentSize = reader.pointSize === 4 || reader.pointSize === 0xfff0 ? 2 : 4;
  if (reader.pointSize !== 4 && reader.pointSize !== 0xfff0 && reader.pointSize !== 8) {
    throw new PptParseError(`Invalid legacy PowerPoint vertex point size ${reader.pointSize}`, {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 325,
      length: reader.pointSize,
      reason: "invalid-freeform-point-size",
    });
  }
  if (reader.offset + componentSize * 2 > reader.view.byteLength) {
    throw new PptParseError("Truncated legacy PowerPoint freeform vertex", {
      part: "record",
      recordType: options.type,
      recordName: "OfficeArtFOPT",
      propertyId: 325,
      byteRange: [reader.offset, reader.offset + componentSize * 2],
      reason: "truncated-freeform-vertex",
    });
  }
  if (componentSize === 4) {
    const x = reader.view.getInt32(reader.offset, true);
    const y = reader.view.getInt32(reader.offset + 4, true);
    reader.offset += 8;
    return { x: String(x), y: String(y) };
  }
  const x = reader.view.getInt16(reader.offset, true);
  const y = reader.view.getInt16(reader.offset + 2, true);
  reader.offset += 4;
  return { x: String(x), y: String(y) };
}

function readRotation(collected: ReadonlyMap<number, CollectedProperty>, style: EscherShapeStyle) {
  const rotation = collected.get(0x0004);
  if (rotation) style.rotation = rotation.value / 65536;
}

function readBlipReference(
  collected: ReadonlyMap<number, CollectedProperty>,
  style: EscherShapeStyle,
) {
  const reference = collected.get(0x0104);
  if (reference?.blip) style.blipReference = reference.value;
}

function readFill(
  collected: ReadonlyMap<number, CollectedProperty>,
  style: EscherShapeStyle,
  options: RecordNode,
) {
  const fillType = collected.get(0x0180);
  const foreground = collected.get(0x0181);
  const opacity = collected.get(0x0182);
  const background = collected.get(0x0183);
  const fillBlip = collected.get(0x0186);
  const angle = collected.get(0x018b);
  if (!fillType) return;
  style.fillType = fillType.value;
  if (fillType.value === 0) {
    style.fill = { type: "none" };
    return;
  }
  if (fillType.value === 1 && foreground) {
    style.fill = { type: "solid", color: solidColor(foreground.value, fillOpacity(opacity)) };
    return;
  }
  if (fillType.value === 5 && fillBlip) {
    style.fillBlipReference = fillBlip.value;
    return;
  }
  if ((fillType.value === 6 || fillType.value === 8) && foreground) {
    const stops: GradientStopOptions[] = [{ position: 0, color: escherColorHex(foreground.value) }];
    if (background) stops.push({ position: 100, color: escherColorHex(background.value) });
    style.fill = {
      type: "gradient",
      ...(angle ? { angle: angle.value / 65536 } : {}),
      stops,
    };
    return;
  }
  throw new PptParseError(`Unsupported legacy PowerPoint fill type ${fillType.value}`, {
    part: "record",
    recordType: options.type,
    recordName: "OfficeArtFOPT",
    propertyId: 0x0180,
    byteRange: [fillType.offset, fillType.offset + 6],
    reason: "unsupported-fill-type",
  });
}

function readOutline(collected: ReadonlyMap<number, CollectedProperty>, style: EscherShapeStyle) {
  const color = collected.get(0x01c0);
  const width = collected.get(0x01cb);
  const dash = collected.get(0x01ce);
  if (!color && !width && !dash && !collected.has(0x01d1) && !collected.has(0x01d2)) return;
  const outline: OutlineOptions = {};
  if (color) {
    outline.type = "solidFill";
    outline.color = escherColorHex(color.value);
  }
  if (width) outline.width = width.value;
  if (dash && LINE_DASHES[dash.value] !== undefined) outline.dash = LINE_DASHES[dash.value];
  const head = lineEnd(collected, 0x01d1, 0x01d3, 0x01d4);
  const tail = lineEnd(collected, 0x01d2, 0x01d5, 0x01d6);
  if (head) outline.headEnd = head;
  if (tail) outline.tailEnd = tail;
  style.outline = outline;
}

function lineEnd(
  collected: ReadonlyMap<number, CollectedProperty>,
  typeId: number,
  widthId: number,
  lengthId: number,
): LineEndOptions | undefined {
  const type = collected.get(typeId);
  if (!type) return undefined;
  const arrowType = ARROW_HEADS[type.value];
  if (!arrowType) return undefined;
  const width = collected.get(widthId);
  const length = collected.get(lengthId);
  return {
    type: arrowType,
    ...(width && ARROW_SIZES[width.value] !== undefined
      ? { width: ARROW_SIZES[width.value] as NonNullable<LineEndOptions["width"]> }
      : {}),
    ...(length && ARROW_SIZES[length.value] !== undefined
      ? { length: ARROW_SIZES[length.value] as NonNullable<LineEndOptions["length"]> }
      : {}),
  };
}

function readShadow(collected: ReadonlyMap<number, CollectedProperty>, style: EscherShapeStyle) {
  const shadowType = collected.get(0x0200);
  if (!shadowType || shadowType.value === 0) return;
  const offsetX = collected.get(0x0205)?.value ?? 0;
  const offsetY = collected.get(0x0206)?.value ?? 0;
  const shadowColor = collected.get(0x0201);
  style.effects = {
    outerShadow: {
      color: { value: shadowColor ? escherColorHex(shadowColor.value) : "000000" },
      distance: Math.round(Math.hypot(offsetX, offsetY)),
      direction: Math.round((Math.atan2(offsetY, offsetX) * 180) / Math.PI),
    },
  };
}

function rejectUnknownProperties(
  collected: ReadonlyMap<number, CollectedProperty>,
  options: RecordNode,
) {
  for (const [id, entry] of collected) {
    if (KNOWN_PROPERTY_IDS.has(id) || isIgnorableProperty(id)) continue;
    throw new PptParseError(
      `Unsupported legacy PowerPoint shape property 0x${id.toString(16).padStart(4, "0")}`,
      {
        part: "record",
        recordType: options.type,
        recordName: "OfficeArtFOPT",
        propertyId: id,
        byteRange: [entry.offset, entry.offset + 6],
        reason: entry.blip ? "unsupported-visual-property" : "unknown-visual-property",
      },
    );
  }
}

function escherColorHex(value: number): string {
  return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff]
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function solidColor(value: number, opacity: RgbColorOptions["transforms"]): RgbColorOptions {
  return {
    value: escherColorHex(value),
    ...(opacity ? { transforms: opacity } : {}),
  };
}

function fillOpacity(entry: CollectedProperty | undefined): RgbColorOptions["transforms"] {
  if (!entry || entry.value >= 65536) return undefined;
  return { alpha: Math.max(0, Math.round((entry.value / 65536) * 100)) };
}

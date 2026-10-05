import type {
  EffectListOptions,
  FillOptions,
  GradientStopOptions,
  LineEndOptions,
  OutlineOptions,
  RgbColorOptions,
  ShapeType,
} from "@office-open/core/drawing";

import { PptParseError } from "../errors";
import type { RecordNode } from "../records";

/** Escher shape-instance → OOXML preset geometry for ordinary autoshapes. */
export const SHAPE_TYPE_PRESETS: Readonly<Record<number, ShapeType>> = {
  1: "rect",
  2: "parallelogram",
  3: "trapezoid",
  4: "diamond",
  5: "roundRect",
  6: "octagon",
  7: "triangle",
  8: "rtTriangle",
  9: "ellipse",
  10: "hexagon",
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
  /** Rotation in degrees (0x0004 is 1/65536-degree fixed point). */
  rotation?: number;
}

interface CollectedProperty {
  readonly value: number;
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
  for (let index = 0; index < options.instance; index += 1) {
    const offset = options.offset + 8 + index * 6;
    const packedId = view.getUint16(offset, true);
    collected.set(packedId & 0x3fff, {
      value: view.getUint32(offset + 2, true),
      blip: (packedId & 0x4000) !== 0,
      offset,
    });
  }
  const style: EscherShapeStyle = {};
  readRotation(collected, style);
  readBlipReference(collected, style);
  readFill(collected, style, options);
  readOutline(collected, style);
  readShadow(collected, style);
  rejectUnknownProperties(collected, options);
  return style;
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

import type { ShapeType } from "@office-open/core/drawing";
import type {
  ConnectorOptions,
  DrawingAnchorOptions,
  PictureOptions,
  ShapeOptions,
  NoteAnchorOptions,
} from "@office-open/xlsx";

const enum EscherRecord {
  Dgg = 0xf000,
  BStoreContainer = 0xf001,
  Dg = 0xf002,
  SpgrContainer = 0xf003,
  SpContainer = 0xf004,
  SolverContainer = 0xf005,
  BlipFirst = 0xf01a,
  BlipLast = 0xf021,
  Bse = 0xf007,
  Blob = 0xf008,
  Sp = 0xf00a,
  Opt = 0xf00b,
  ClientAnchor = 0xf010,
  ClientData = 0xf011,
}

interface EscherRecordHeader {
  instance: number;
  type: number;
  body: Uint8Array;
}

interface EscherAnchor {
  fromColumn: number;
  fromRow: number;
  fromColumnOffset: number;
  fromRowOffset: number;
  toColumn: number;
  toRow: number;
  toColumnOffset: number;
  toRowOffset: number;
}

interface EscherDrawing {
  anchor?: EscherAnchor;
  shapeId?: number;
  instance?: number;
  connector?: boolean;
  blipReference?: number;
  clientData?: boolean;
}

export interface EscherImage {
  data: Uint8Array;
  type: "png" | "jpg" | "wmf" | "emf";
}

export interface EscherDrawings {
  connectors: ConnectorOptions[];
  images: PictureOptions[];
  shapes: ShapeOptions[];
}

export interface EscherAnchorProjection {
  anchor: NoteAnchorOptions;
  shapeId?: number;
}

interface ParsedEscher {
  drawings: EscherDrawing[];
  images: EscherImage[];
}

// Exact equivalents only: MS-ODRAW MSOSPT -> DrawingML ST_ShapeType.
// One-dimensional, picture-frame, text-box, WordArt, and host-control
// instances are handled by their canonical projection or rejected explicitly.
const SHAPE_PRESETS: Readonly<Record<number, ShapeType>> = {
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
  16: "cube",
  19: "arc",
  20: "line",
  21: "plaque",
  22: "can",
  23: "donut",
  41: "callout1",
  42: "callout2",
  43: "callout3",
  44: "accentCallout1",
  45: "accentCallout2",
  46: "accentCallout3",
  47: "borderCallout1",
  48: "borderCallout2",
  49: "borderCallout3",
  50: "accentBorderCallout1",
  51: "accentBorderCallout2",
  52: "accentBorderCallout3",
  53: "ribbon",
  54: "ribbon2",
  55: "chevron",
  56: "pentagon",
  57: "noSmoking",
  58: "star8",
  59: "star16",
  60: "star32",
  61: "wedgeRectCallout",
  62: "wedgeRoundRectCallout",
  63: "wedgeEllipseCallout",
  64: "wave",
  65: "foldedCorner",
  66: "leftArrow",
  67: "downArrow",
  68: "upArrow",
  69: "leftRightArrow",
  70: "upDownArrow",
  71: "irregularSeal1",
  72: "irregularSeal2",
  73: "lightningBolt",
  74: "heart",
  76: "quadArrow",
  77: "leftArrowCallout",
  78: "rightArrowCallout",
  79: "upArrowCallout",
  80: "downArrowCallout",
  81: "leftRightArrowCallout",
  82: "upDownArrowCallout",
  83: "quadArrowCallout",
  84: "bevel",
  85: "leftBracket",
  86: "rightBracket",
  87: "leftBrace",
  88: "rightBrace",
  89: "leftUpArrow",
  90: "bentUpArrow",
  91: "bentArrow",
  92: "star24",
  93: "stripedRightArrow",
  94: "notchedRightArrow",
  95: "blockArc",
  96: "smileyFace",
  97: "verticalScroll",
  98: "horizontalScroll",
  99: "circularArrow",
  101: "uturnArrow",
  102: "curvedRightArrow",
  103: "curvedLeftArrow",
  104: "curvedUpArrow",
  105: "curvedDownArrow",
  106: "cloudCallout",
  107: "ellipseRibbon",
  108: "ellipseRibbon2",
  109: "flowChartProcess",
  110: "flowChartDecision",
  111: "flowChartInputOutput",
  112: "flowChartPredefinedProcess",
  113: "flowChartInternalStorage",
  114: "flowChartDocument",
  115: "flowChartMultidocument",
  116: "flowChartTerminator",
  117: "flowChartPreparation",
  118: "flowChartManualInput",
  119: "flowChartManualOperation",
  120: "flowChartConnector",
  121: "flowChartPunchedCard",
  122: "flowChartPunchedTape",
  123: "flowChartSummingJunction",
  124: "flowChartOr",
  125: "flowChartCollate",
  126: "flowChartSort",
  127: "flowChartExtract",
  128: "flowChartMerge",
  129: "flowChartOfflineStorage",
  130: "flowChartOnlineStorage",
  131: "flowChartMagneticTape",
  132: "flowChartMagneticDisk",
  133: "flowChartMagneticDrum",
  134: "flowChartDisplay",
  135: "flowChartDelay",
  176: "flowChartAlternateProcess",
  177: "flowChartOffpageConnector",
  182: "leftRightUpArrow",
  183: "sun",
  184: "moon",
  185: "bracketPair",
  186: "bracePair",
  187: "star4",
  188: "doubleWave",
  189: "actionButtonBlank",
  190: "actionButtonHome",
  191: "actionButtonHelp",
  192: "actionButtonInformation",
  193: "actionButtonForwardNext",
  194: "actionButtonBackPrevious",
  195: "actionButtonEnd",
  196: "actionButtonBeginning",
  197: "actionButtonReturn",
  198: "actionButtonDocument",
  199: "actionButtonSound",
  200: "actionButtonMovie",
};

export const CONNECTOR_PRESETS: Readonly<Record<number, ShapeType>> = {
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

function imageType(bytes: Uint8Array): EscherImage["type"] {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8) return "jpg";
  if (bytes.length >= 4 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e)
    return "png";
  if (bytes.length >= 4 && bytes[0] === 0x01 && bytes[1] === 0x00 && bytes[2] === 0x00)
    return "emf";
  return "wmf";
}

function recordAt(data: Uint8Array, position: number): EscherRecordHeader | undefined {
  if (position < 0 || position + 8 > data.byteLength) return undefined;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const versionInstance = view.getUint16(position, true);
  const type = view.getUint16(position + 2, true);
  const length = view.getUint32(position + 4, true);
  const start = position + 8;
  if (length > data.byteLength - start) return undefined;
  return {
    instance: versionInstance >> 4,
    type,
    body: data.subarray(start, start + length),
  };
}

function isContainer(type: number): boolean {
  return (
    type === EscherRecord.Dgg ||
    type === EscherRecord.Dg ||
    type === EscherRecord.BStoreContainer ||
    type === EscherRecord.SpgrContainer ||
    type === EscherRecord.SolverContainer ||
    type === EscherRecord.SpContainer
  );
}

function isBlip(type: number, inBstore: boolean): boolean {
  if (
    type === EscherRecord.Sp ||
    type === EscherRecord.Opt ||
    type === EscherRecord.ClientAnchor ||
    type === EscherRecord.ClientData
  ) {
    return false;
  }
  return inBstore && type >= EscherRecord.BlipFirst && type <= EscherRecord.BlipLast;
}

function parseAnchor(body: Uint8Array): EscherAnchor {
  if (body.byteLength < 18)
    throw new XlsParseError("Invalid legacy XLS file: truncated Escher anchor");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  return {
    fromRow: view.getInt16(2, true),
    fromColumn: view.getInt16(4, true),
    fromRowOffset: view.getInt32(6, true),
    fromColumnOffset: view.getInt32(10, true),
    toRow: view.getInt16(14, true),
    toColumn: view.getInt16(16, true),
    toRowOffset: 0,
    toColumnOffset: 0,
  };
}

function readAnchorTail(body: Uint8Array, start: number, anchor: EscherAnchor): void {
  if (body.byteLength < start + 12) return;
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  anchor.toRowOffset = view.getInt32(start, true);
  anchor.toColumnOffset = view.getInt32(start + 4, true);
  anchor.toRow = view.getInt16(start + 8, true);
  anchor.toColumn = view.getInt16(start + 10, true);
}

function parseContainer(
  data: Uint8Array,
  position: number,
  limit: number,
  result: ParsedEscher,
  inBstore = false,
  drawing?: EscherDrawing,
): number {
  let cursor = position;
  while (cursor + 8 <= limit) {
    const record = recordAt(data, cursor);
    if (!record) return limit;
    const headerEnd = cursor + 8;
    const bodyEnd = headerEnd + record.body.byteLength;
    if (isContainer(record.type)) {
      const shapeContainer = record.type === EscherRecord.SpContainer;
      const childDrawing = shapeContainer ? {} : drawing;
      parseContainer(
        data,
        headerEnd,
        bodyEnd,
        result,
        record.type === EscherRecord.BStoreContainer,
        childDrawing,
      );
      if (shapeContainer && childDrawing?.anchor) result.drawings.push(childDrawing);
    } else if (record.type === EscherRecord.Bse) {
      const blip = recordAt(data, headerEnd + 36);
      if (blip && isBlip(blip.type, true))
        result.images.push({ data: blip.body, type: imageType(blip.body) });
    } else if (isBlip(record.type, inBstore)) {
      result.images.push({ data: record.body, type: imageType(record.body) });
    } else if (record.type === EscherRecord.Sp) {
      if (!drawing) throw new XlsParseError("Invalid legacy XLS file: orphan Escher shape");
      if (record.body.byteLength < 8)
        throw new XlsParseError("Invalid legacy XLS file: truncated Escher shape");
      const view = new DataView(record.body.buffer, record.body.byteOffset, record.body.byteLength);
      drawing.shapeId = view.getInt32(0, true) & 0x000fffff;
      drawing.instance = record.instance;
      drawing.connector = (view.getUint32(4, true) & 0x0100) !== 0;
    } else if (record.type === EscherRecord.Opt) {
      if (drawing) {
        for (let offset = 0; offset + 6 <= record.body.byteLength; offset += 6) {
          const view = new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          );
          if ((view.getUint16(offset, true) & 0x3fff) === 0x0104)
            drawing.blipReference = view.getUint16(offset + 4, true);
        }
      }
    } else if (record.type === EscherRecord.ClientData) {
      if (!drawing) throw new XlsParseError("Invalid legacy XLS file: orphan Escher client data");
      drawing.clientData = true;
    } else if (record.type === EscherRecord.ClientAnchor) {
      if (!drawing) throw new XlsParseError("Invalid legacy XLS file: orphan Escher anchor");
      drawing.anchor = parseAnchor(record.body);
      readAnchorTail(record.body, 18, drawing.anchor);
    }
    cursor = bodyEnd;
  }
  return cursor;
}

export function parseEscher(data: Uint8Array): ParsedEscher {
  const result: ParsedEscher = { drawings: [], images: [] };
  parseContainer(data, 0, data.byteLength, result);
  return result;
}

function anchorOptions(anchor: EscherAnchor): DrawingAnchorOptions {
  return {
    col: Math.max(0, anchor.fromColumn) + 1,
    row: Math.max(0, anchor.fromRow) + 1,
    colOffset: Math.max(0, anchor.fromColumnOffset),
    rowOffset: Math.max(0, anchor.fromRowOffset),
    toCol: Math.max(0, anchor.toColumn) + 1,
    toRow: Math.max(0, anchor.toRow) + 1,
    toColOffset: Math.max(0, anchor.toColumnOffset),
    toRowOffset: Math.max(0, anchor.toRowOffset),
    anchorType: "twoCell" as const,
  };
}

function noteAnchor(anchor: EscherAnchor): NoteAnchorOptions {
  return {
    from: {
      col: Math.max(0, anchor.fromColumn),
      row: Math.max(0, anchor.fromRow),
      colOff: Math.max(0, anchor.fromColumnOffset),
      rowOff: Math.max(0, anchor.fromRowOffset),
    },
    to: {
      col: Math.max(0, anchor.toColumn),
      row: Math.max(0, anchor.toRow),
      colOff: Math.max(0, anchor.toColumnOffset),
      rowOff: Math.max(0, anchor.toRowOffset),
    },
  };
}

function drawingImage(
  drawing: EscherDrawing,
  workbookImages: readonly EscherImage[],
  fallbackImages: readonly EscherImage[],
): EscherImage {
  const index = drawing.blipReference ?? -1;
  const image = workbookImages[index] ?? fallbackImages[index] ?? fallbackImages[0];
  if (!image) throw new XlsParseError("Invalid legacy XLS file: missing Escher picture stream");
  return image;
}

function unsupportedShape(instance: number): XlsParseError {
  return new XlsParseError(
    `Invalid legacy XLS file: unsupported Escher shape instance ${instance}`,
    {
      part: "data",
      recordName: "OfficeArtFSP",
      recordType: 0xf00a,
      reason: "unsupported-visual-shape",
    },
  );
}

export function escherDrawings(
  data: Uint8Array,
  workbookImages: EscherImage[] = [],
): EscherDrawings {
  const parsed = parseEscher(data);
  const drawings: EscherDrawings = { connectors: [], images: [], shapes: [] };
  const anchored = parsed.drawings.filter((drawing) => drawing.anchor && !drawing.clientData);
  for (const [index, drawing] of anchored.entries()) {
    const anchor = anchorOptions(drawing.anchor!);
    if (drawing.instance === 75 && drawing.blipReference === undefined)
      throw new XlsParseError("Invalid legacy XLS file: missing Escher picture stream", {
        part: "data",
        recordName: "OfficeArtFSP",
        recordType: 0xf00a,
        reason: "missing-picture-reference",
      });
    const picture = drawing.blipReference !== undefined;
    if (picture) {
      const image = drawingImage(drawing, workbookImages, parsed.images);
      drawings.images.push({
        ...anchor,
        name: `Picture ${index + 1}`,
        shapeId: drawing.shapeId,
        ...image,
      });
    } else if (drawing.connector) {
      const preset =
        drawing.instance === undefined ? undefined : CONNECTOR_PRESETS[drawing.instance];
      if (!preset) throw unsupportedShape(drawing.instance ?? 0);
      drawings.connectors.push({
        ...anchor,
        name: `Connector ${index + 1}`,
        shapeId: drawing.shapeId,
        properties: { geometry: preset },
      });
    } else {
      if (drawing.instance === 202) {
        drawings.shapes.push({
          ...anchor,
          name: `Text Box ${index + 1}`,
          shapeId: drawing.shapeId,
          textBox: true,
          properties: { geometry: "rect" },
        });
      } else {
        const preset = drawing.instance === undefined ? undefined : SHAPE_PRESETS[drawing.instance];
        if (!preset) throw unsupportedShape(drawing.instance ?? 0);
        drawings.shapes.push({
          ...anchor,
          name: `Shape ${index + 1}`,
          shapeId: drawing.shapeId,
          properties: { geometry: preset },
        });
      }
    }
  }
  return drawings;
}

export function escherAnchors(data: Uint8Array): EscherAnchorProjection[] {
  return parseEscher(data)
    .drawings.filter((drawing) => drawing.anchor)
    .map((drawing) => ({ anchor: noteAnchor(drawing.anchor!), shapeId: drawing.shapeId }));
}

export function escherPictures(
  data: Uint8Array,
  workbookImages: EscherImage[] = [],
): PictureOptions[] {
  return escherDrawings(data, workbookImages).images;
}

export function escherImages(data: Uint8Array): EscherImage[] {
  const images = parseEscher(data);
  return images.images;
}
import { XlsParseError } from "../errors";

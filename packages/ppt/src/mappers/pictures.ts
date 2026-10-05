import { createView } from "../parts/persist";
import {
  LegacyPowerPointError,
  isZeroPadding,
  readRecordHeader,
  recordBody,
  RecordType,
  type RecordHeader,
  type RecordNode,
} from "../records";
import { PICTURE_RECORD_TYPES, RECORD_BODY_OFFSET } from "./constants";
import type { LegacyPicture } from "./models";
export function readPictureStore(data: Uint8Array): readonly (LegacyPicture | undefined)[] {
  const view = createView(data);
  const pictureRecords = readPictureRecords(view);
  return [undefined, ...pictureRecords.map((record) => readIsolatedPicture(view, record))];
}

function readPictureRecords(view: DataView): readonly RecordNode[] {
  const records = readPictureRecordRange(view, 0, view.byteLength);
  return records.length === 1 && records[0]!.type === RecordType.escherBStoreContainer
    ? readPictureRecordRange(view, records[0]!.offset + RECORD_BODY_OFFSET, records[0]!.end)
    : records;
}

function readPictureRecordRange(view: DataView, offset: number, end: number): RecordNode[] {
  const records: RecordNode[] = [];
  let cursor = offset;
  while (cursor + RECORD_BODY_OFFSET <= end) {
    if (isZeroPadding(view, cursor, end)) break;
    let header: RecordHeader;
    try {
      header = readRecordHeader(view, cursor, end);
    } catch (error) {
      if (!(error instanceof LegacyPowerPointError)) throw error;
      break;
    }
    records.push({ ...header, children: [] });
    if (header.end <= cursor) break;
    cursor = header.end;
  }
  return records;
}

function readIsolatedPicture(view: DataView, record: RecordNode): LegacyPicture | undefined {
  if (!isPictureRecord(record)) return undefined;
  try {
    return record.type === RecordType.escherBse
      ? readPictureRecord(view, record)
      : readEmbeddedPicture(record, recordBody(view, record));
  } catch (error) {
    if (!(error instanceof LegacyPowerPointError)) throw error;
    return undefined;
  }
}

function isPictureRecord(record: RecordNode): boolean {
  return record.type === RecordType.escherBse || PICTURE_RECORD_TYPES.has(record.type);
}

function readPictureRecord(view: DataView, bse: RecordNode): LegacyPicture | undefined {
  if (bse.length < 36) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed Pictures stream");
  }

  const nameLength = view.getUint8(bse.offset + 8 + 33);
  const blipOffset = bse.offset + 8 + 36 + nameLength;
  if (blipOffset === bse.end) return undefined;
  if (blipOffset + 8 > bse.end) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed Pictures stream");
  }

  const blip = readRecordHeader(view, blipOffset, bse.end);
  if (blip.end !== bse.end) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed Pictures stream");
  }
  const body = recordBody(view, blip);
  return readEmbeddedPicture(blip, body);
}

function readEmbeddedPicture(blip: RecordHeader, body: Uint8Array): LegacyPicture | undefined {
  const skip = picturePayloadOffset(blip);
  if (skip === undefined) return undefined;
  if (skip > body.byteLength) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed picture data");
  }
  const data = body.slice(skip);
  if (blip.type === 0xf01f) return { type: "bmp", data: bitmapFromDib(data) };
  return { type: pictureType(blip.type), data };
}

function picturePayloadOffset(blip: RecordHeader): number | undefined {
  if (blip.type === 0xf01a) {
    if (blip.instance === 0x3d4) return 50;
    if (blip.instance === 0x3d5) return 66;
  } else if (blip.type === 0xf01b) {
    if (blip.instance === 0x216) return 50;
    if (blip.instance === 0x217) return 66;
  } else if (blip.type === 0xf01d || blip.type === 0xf02a) {
    if (blip.instance === 0x46a || blip.instance === 0x6e2) return 17;
    if (blip.instance === 0x46b || blip.instance === 0x6e3) return 33;
  } else if (blip.type === 0xf01e) {
    if (blip.instance === 0x6e0) return 17;
    if (blip.instance === 0x6e1) return 33;
  } else if (blip.type === 0xf01f) {
    if (blip.instance === 0x7a8) return 17;
    if (blip.instance === 0x7a9) return 33;
  } else if (blip.type === 0xf01c || blip.type === 0xf029) {
    return undefined;
  } else {
    return undefined;
  }
  throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: unsupported picture variant");
}

function pictureType(type: number): LegacyPicture["type"] {
  if (type === 0xf01a) return "emf";
  if (type === 0xf01b) return "wmf";
  if (type === 0xf01d || type === 0xf02a) return "jpg";
  if (type === 0xf01e) return "png";
  return "bmp";
}

function bitmapFromDib(dib: Uint8Array): Uint8Array {
  if (dib.byteLength < 40) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed picture data");
  }
  const view = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const headerSize = view.getUint32(0, true);
  const compression = view.getUint32(16, true);
  const colorsUsed = view.getUint32(32, true);
  const bitCount = view.getUint16(14, true);
  const paletteSize = colorsUsed > 0 ? colorsUsed * 4 : bitCount <= 8 ? (1 << bitCount) * 4 : 0;
  const maskSize = compression === 3 && headerSize === 40 ? 12 : 0;
  const pixelOffset = 14 + headerSize + paletteSize + maskSize;
  if (
    headerSize < 40 ||
    headerSize > dib.byteLength ||
    (compression !== 0 && compression !== 3) ||
    pixelOffset > dib.byteLength
  ) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed picture data");
  }
  const bitmap = new Uint8Array(14 + dib.byteLength);
  const output = new DataView(bitmap.buffer);
  output.setUint16(0, 0x424d, true);
  output.setUint32(2, bitmap.byteLength, true);
  output.setUint32(10, pixelOffset, true);
  bitmap.set(dib, 14);
  return bitmap;
}

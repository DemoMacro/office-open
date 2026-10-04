import type { PictureOptions } from "@office-open/xlsx";

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

export interface EscherImage {
  data: Uint8Array;
  type: PictureOptions["type"];
}

interface ParsedEscher {
  anchors: EscherAnchor[];
  images: EscherImage[];
  shapeIds: number[];
  blipReferences: number[];
}

function imageType(bytes: Uint8Array): PictureOptions["type"] {
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
    throw new LegacyExcelError("Invalid legacy XLS file: truncated Escher anchor");
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
): number {
  let cursor = position;
  let currentAnchor: EscherAnchor | undefined;
  while (cursor + 8 <= limit) {
    const record = recordAt(data, cursor);
    if (!record) return limit;
    const headerEnd = cursor + 8;
    const bodyEnd = headerEnd + record.body.byteLength;
    if (isContainer(record.type)) {
      parseContainer(
        data,
        headerEnd,
        bodyEnd,
        result,
        record.type === EscherRecord.BStoreContainer,
      );
    } else if (record.type === EscherRecord.Bse) {
      const blip = recordAt(data, headerEnd + 36);
      if (blip && isBlip(blip.type, true))
        result.images.push({ data: blip.body, type: imageType(blip.body) });
    } else if (isBlip(record.type, inBstore)) {
      result.images.push({ data: record.body, type: imageType(record.body) });
    } else if (record.type === EscherRecord.Sp) {
      if (record.body.byteLength >= 8) {
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        result.shapeIds.push(view.getInt32(4, true) & 0x000fffff);
      }
    } else if (record.type === EscherRecord.Opt) {
      for (let offset = 0; offset + 6 <= record.body.byteLength; offset += 6) {
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const property = view.getUint16(offset, true);
        if (property === 0x0104) result.blipReferences.push(view.getUint16(offset + 4, true));
      }
    } else if (record.type === EscherRecord.ClientAnchor) {
      currentAnchor = parseAnchor(record.body);
      readAnchorTail(record.body, 18, currentAnchor);
      result.anchors.push(currentAnchor);
    }
    cursor = bodyEnd;
  }
  return cursor;
}

export function parseEscher(data: Uint8Array): ParsedEscher {
  const result: ParsedEscher = { anchors: [], images: [], shapeIds: [], blipReferences: [] };
  parseContainer(data, 0, data.byteLength, result);
  return result;
}

export function escherPictures(
  data: Uint8Array,
  workbookImages: EscherImage[] = [],
): PictureOptions[] {
  const parsed = parseEscher(data);
  return parsed.anchors.map((anchor, index) => ({
    name: `Picture ${index + 1}`,
    type:
      workbookImages[parsed.blipReferences[index] ?? -1]?.type ??
      parsed.images[index]?.type ??
      parsed.images[0]?.type ??
      "png",
    data:
      workbookImages[parsed.blipReferences[index] ?? -1]?.data ??
      parsed.images[index]?.data ??
      parsed.images[0]?.data ??
      new Uint8Array(),
    col: Math.max(0, anchor.fromColumn) + 1,
    row: Math.max(0, anchor.fromRow) + 1,
    colOffset: Math.max(0, anchor.fromColumnOffset),
    rowOffset: Math.max(0, anchor.fromRowOffset),
    toCol: Math.max(0, anchor.toColumn) + 1,
    toRow: Math.max(0, anchor.toRow) + 1,
    toColOffset: Math.max(0, anchor.toColumnOffset),
    toRowOffset: Math.max(0, anchor.toRowOffset),
    anchorType: "twoCell" as const,
    shapeId: parsed.shapeIds[index],
  }));
}

export function escherImages(data: Uint8Array): EscherImage[] {
  const images = parseEscher(data);
  return images.images;
}
import { LegacyExcelError } from "./errors";

import { DocParseError } from "../errors";

export interface NumberPair {
  offset: number;
  length: number;
}

export interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
}

const UNICODE_PIECE = 0x40000000;
const PIECE_OFFSET_MASK = 0x3fffffff;

export function transformPieceFc(value: number): { fc: number; compressed: boolean } {
  const compressed = (value & UNICODE_PIECE) !== 0;
  return {
    compressed,
    fc: compressed ? (value & PIECE_OFFSET_MASK) / 2 : value & PIECE_OFFSET_MASK,
  };
}

export function parsePieceTable(table: Uint8Array, clx: NumberPair): Piece[] {
  const rangeMessage = "Invalid DOC CLX: piece table is outside the table stream";
  requireRange(table, clx.offset, clx.length, rangeMessage);
  if (clx.length === 0)
    throw new DocParseError("Invalid Word document: CLX piece table is missing");

  let position = clx.offset;
  const end = clx.offset + clx.length;
  while (position < end) {
    const kind = table[position]!;
    if (kind === 1) {
      requireRange(table, position, 3, "Invalid DOC CLX: truncated property record");
      position += 3 + table[position + 1]! + (table[position + 2]! << 8);
      continue;
    }
    if (kind === 2) {
      requireRange(table, position, 5, "Invalid DOC CLX: truncated piece table header");
      const length = readUint32(table, position + 1, "Invalid DOC CLX: invalid piece table length");
      const tableStart = position + 5;
      requireRange(table, tableStart, length, rangeMessage);
      if (tableStart + length !== end) {
        throw new DocParseError("Invalid DOC CLX: piece table does not terminate CLX");
      }
      if (length < 12 || (length - 4) % 12 !== 0) {
        throw new DocParseError("Invalid DOC CLX: malformed piece table size");
      }
      return parsePieces(table, tableStart, length);
    }
    throw new DocParseError(`Invalid DOC CLX: unknown record type ${kind}`);
  }
  throw new DocParseError("Invalid DOC CLX: piece table marker is missing");
}

function parsePieces(table: Uint8Array, start: number, length: number): Piece[] {
  const pieceCount = (length - 4) / 12;
  const pieces: Piece[] = [];
  for (let index = 0; index < pieceCount; index++) {
    const cpOffset = start + index * 4;
    const cpStart = readUint32(
      table,
      cpOffset,
      "Invalid DOC piece table: CP is outside the piece table",
    );
    const cpEnd = readUint32(
      table,
      cpOffset + 4,
      "Invalid DOC piece table: CP is outside the piece table",
    );
    if (index === 0 && cpStart !== 0) {
      throw new DocParseError("Invalid DOC piece table: first CP is not zero");
    }
    if (cpEnd <= cpStart || (index > 0 && cpStart !== pieces[index - 1]!.cpEnd)) {
      throw new DocParseError("Invalid DOC piece table: CP boundaries are not contiguous");
    }

    const pcdOffset = start + (pieceCount + 1) * 4 + index * 8;
    const fcValue = readUint32(
      table,
      pcdOffset + 2,
      "Invalid DOC piece table: PCD is outside the piece table",
    );
    const { fc, compressed } = transformPieceFc(fcValue);
    if (!Number.isSafeInteger(fc) || fc < 0) {
      throw new DocParseError("Invalid DOC piece table: invalid compressed FC");
    }
    pieces.push({ cpStart, cpEnd, fc, compressed });
  }
  return pieces;
}

function requireRange(bytes: Uint8Array, offset: number, length: number, message: string): void {
  if (offset < 0 || length < 0 || offset > bytes.byteLength || length > bytes.byteLength - offset) {
    throw new DocParseError(message);
  }
}

function readUint32(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 4, message);
  return (
    (bytes[offset]! |
      (bytes[offset + 1]! << 8) |
      (bytes[offset + 2]! << 16) |
      (bytes[offset + 3]! << 24)) >>>
    0
  );
}

import { DocParseError } from "../errors";
import type { Piece } from "../streams/pieces";
import { readUint32, requireRange } from "./binary";
import { FKP_LENGTH } from "./constants";
import type { CharacterRange, ParagraphRange } from "./models";
import { parseParagraphProperties } from "./paragraph-formatting";
import { parseCharacterProperties } from "./run-formatting";
export function fcToCp(fc: number, piece: Piece): number {
  const characters = piece.compressed ? fc - piece.fc : (fc - piece.fc) / 2;
  if (!Number.isInteger(characters)) {
    throw new DocParseError("Invalid DOC property range: FC is not aligned to the text piece");
  }
  return piece.cpStart + characters;
}

export function overlappingPieces(
  pieces: Piece[],
  cpEndLimit: number,
  fcStart: number,
  fcEnd: number,
): Array<{ cpStart: number; cpEnd: number; piece: Piece }> {
  const result: Array<{ cpStart: number; cpEnd: number; piece: Piece }> = [];
  for (const piece of pieces) {
    const cpEnd = Math.min(piece.cpEnd, cpEndLimit);
    if (piece.cpStart >= cpEnd) continue;
    const pieceFcEnd = piece.fc + (cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
    if (piece.fc >= fcEnd || pieceFcEnd <= fcStart) continue;
    const cpIntersectionStart = Math.max(piece.cpStart, fcToCp(Math.max(fcStart, piece.fc), piece));
    const cpIntersectionEnd = Math.min(cpEnd, fcToCp(Math.min(fcEnd, pieceFcEnd), piece));
    if (cpIntersectionEnd > cpIntersectionStart) {
      result.push({ cpStart: cpIntersectionStart, cpEnd: cpIntersectionEnd, piece });
    }
  }
  return result;
}

export function parseCharacterFkp(
  word: Uint8Array,
  page: number,
  pieces: Piece[],
  cpEndLimit: number,
): CharacterRange[] {
  requireRange(
    word,
    page,
    FKP_LENGTH,
    "Invalid DOC character FKP: page is outside the WordDocument stream",
  );
  const count = word[page + FKP_LENGTH - 1]!;
  if (count === 0 || count > (FKP_LENGTH - 5) / 5) {
    throw new DocParseError("Invalid DOC character FKP: invalid entry count");
  }
  const ranges: CharacterRange[] = [];
  for (let index = 0; index < count; index++) {
    const fcStart = readUint32(
      word,
      page + index * 4,
      "Invalid DOC character FKP: FC array is truncated",
    );
    const fcEnd = readUint32(
      word,
      page + (index + 1) * 4,
      "Invalid DOC character FKP: FC array is truncated",
    );
    const propertyOffset = word[page + (count + 1) * 4 + index]!;
    if (fcEnd <= fcStart) throw new DocParseError("Invalid DOC character FKP: invalid FC range");
    if (propertyOffset === 0) continue;
    const propertiesStart = page + propertyOffset * 2;
    const propertiesLength = word[propertiesStart]!;
    const grpprlStart = propertiesStart + 1;
    requireRange(
      word,
      grpprlStart,
      propertiesLength,
      "Invalid DOC character FKP: property data is outside the page",
    );
    const properties = parseCharacterProperties(
      word.subarray(grpprlStart, grpprlStart + propertiesLength),
    );
    for (const overlap of overlappingPieces(pieces, cpEndLimit, fcStart, fcEnd)) {
      ranges.push({ ...overlap, properties });
    }
  }
  return ranges.sort((left, right) => left.cpStart - right.cpStart || left.cpEnd - right.cpEnd);
}

export function parseParagraphFkp(word: Uint8Array, page: number): ParagraphRange[] {
  requireRange(
    word,
    page,
    FKP_LENGTH,
    "Invalid DOC paragraph FKP: page is outside the WordDocument stream",
  );
  const count = word[page + FKP_LENGTH - 1]!;
  if (count === 0 || count > Math.floor((FKP_LENGTH - 5) / 17)) {
    throw new DocParseError("Invalid DOC paragraph FKP: invalid entry count");
  }
  const ranges: ParagraphRange[] = [];
  for (let index = 0; index < count; index++) {
    const fcStart = readUint32(
      word,
      page + index * 4,
      "Invalid DOC paragraph FKP: FC array is truncated",
    );
    const fcEnd = readUint32(
      word,
      page + (index + 1) * 4,
      "Invalid DOC paragraph FKP: FC array is truncated",
    );
    const entryOffset = page + (count + 1) * 4 + index * 13;
    const propertyOffset = word[entryOffset]!;
    if (fcEnd <= fcStart) throw new DocParseError("Invalid DOC paragraph FKP: invalid FC range");
    if (propertyOffset === 0) {
      ranges.push({ cpStart: fcStart, cpEnd: fcEnd, properties: {} });
      continue;
    }
    const propertiesStart = page + propertyOffset * 2;
    const wordCount = word[propertiesStart]!;
    const grpprlStart = propertiesStart + (wordCount === 0 ? 4 : 3);
    const propertiesLength =
      wordCount === 0 ? word[propertiesStart + 1]! * 2 - 2 : wordCount * 2 - 3;
    requireRange(
      word,
      grpprlStart,
      propertiesLength,
      "Invalid DOC paragraph FKP: property data is outside the page",
    );
    const properties = parseParagraphProperties(
      word.subarray(grpprlStart, grpprlStart + propertiesLength),
    );
    ranges.push({ cpStart: fcStart, cpEnd: fcEnd, properties });
  }
  return ranges;
}

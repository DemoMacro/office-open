import type { Piece } from "../streams/pieces";
import { requireRange } from "./binary";
import type { TextCharacter } from "./models";
export const WINDOWS_1252_HIGH = new Map<number, string>([
  [0x80, "\u20ac"],
  [0x82, "\u201a"],
  [0x83, "\u0192"],
  [0x84, "\u201e"],
  [0x85, "\u2026"],
  [0x86, "\u2020"],
  [0x87, "\u2021"],
  [0x88, "\u02c6"],
  [0x89, "\u2030"],
  [0x8a, "\u0160"],
  [0x8b, "\u2039"],
  [0x8c, "\u0152"],
  [0x8e, "\u017d"],
  [0x91, "\u2018"],
  [0x92, "\u2019"],
  [0x93, "\u201c"],
  [0x94, "\u201d"],
  [0x95, "\u2022"],
  [0x96, "\u2013"],
  [0x97, "\u2014"],
  [0x98, "\u02dc"],
  [0x99, "\u2122"],
  [0x9a, "\u0161"],
  [0x9b, "\u203a"],
  [0x9c, "\u0153"],
  [0x9e, "\u017e"],
  [0x9f, "\u0178"],
]);

export function decodePiece(word: Uint8Array, piece: Piece): string {
  const byteLength = (piece.cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
  requireRange(
    word,
    piece.fc,
    byteLength,
    "Invalid DOC piece: text is outside the WordDocument stream",
  );
  let result = "";
  if (piece.compressed) {
    for (let index = 0; index < byteLength; index++) {
      const byte = word[piece.fc + index]!;
      result += WINDOWS_1252_HIGH.get(byte) ?? String.fromCharCode(byte);
    }
    return result;
  }
  for (let index = 0; index < byteLength; index += 2) {
    result += String.fromCharCode(word[piece.fc + index]! | (word[piece.fc + index + 1]! << 8));
  }
  return result;
}

function normalizeCharacter(value: string): string | undefined {
  switch (value) {
    case "\r":
    case "\x07":
    case "\x0c":
      return value;
    case "\x1e":
      return "\u2011";
    case "\x1f":
    case "\x01":
    case "\x02":
    case "\x05":
    case "\x08":
    case "\x13":
    case "\x14":
    case "\x15":
      return undefined;
    default:
      return value >= " " ? value : undefined;
  }
}

export function projectText(pieces: Piece[], word: Uint8Array, ccpText: number): TextCharacter[] {
  const characters: TextCharacter[] = [];
  const fieldResults: boolean[] = [];
  for (const piece of pieces) {
    if (piece.cpStart >= ccpText) break;
    const cpEnd = Math.min(piece.cpEnd, ccpText);
    const decoded = decodePiece(word, { ...piece, cpEnd });
    for (let index = 0; index < decoded.length; index++) {
      const source = decoded[index]!;
      if (source === "\x13") {
        fieldResults.push(false);
        continue;
      }
      if (source === "\x14") {
        if (fieldResults.length > 0) fieldResults[fieldResults.length - 1] = true;
        continue;
      }
      if (source === "\x15") {
        fieldResults.pop();
        continue;
      }
      if (fieldResults.at(-1) === false) continue;
      const value = normalizeCharacter(source);
      const isParagraphEnd = value === "\r" || value === "\x07" || value === "\x0c";
      if (isParagraphEnd) {
        characters.push({ value: "", cp: piece.cpStart + index, paragraphEnd: true });
      } else if (value !== undefined) {
        characters.push({ value, cp: piece.cpStart + index });
      }
    }
  }
  return characters;
}

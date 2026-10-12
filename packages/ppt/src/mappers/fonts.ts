import { PptParseError } from "../errors";
import { RecordType, collectDescendants, recordBody, type RecordNode } from "../records";

function decodeNullTerminatedUtf16(bytes: Uint8Array): string {
  let result = "";
  for (let offset = 0; offset + 1 < bytes.byteLength; offset += 2) {
    const character = bytes[offset]! | (bytes[offset + 1]! << 8);
    if (character === 0) return result;
    result += String.fromCharCode(character);
  }
  return result;
}

export function readFontNames(view: DataView, document: RecordNode): ReadonlyMap<number, string> {
  const fonts = new Map<number, string>();
  for (const collection of collectDescendants(document, RecordType.fontCollection)) {
    for (const entity of collection.children) {
      if (entity.type !== RecordType.fontEntityAtom) continue;
      if (entity.length < 68) {
        throw new PptParseError("Corrupt legacy PowerPoint file: truncated font entity", {
          part: "record",
          recordType: entity.type,
          recordName: "FontEntityAtom",
          offset: entity.offset,
          length: entity.length,
          reason: "invalid-record-length",
        });
      }
      const name = decodeNullTerminatedUtf16(recordBody(view, entity).subarray(0, 64));
      if (!name) {
        throw new PptParseError("Corrupt legacy PowerPoint file: missing font name", {
          part: "record",
          recordType: entity.type,
          recordName: "FontEntityAtom",
          offset: entity.offset,
          reason: "missing-font-name",
        });
      }
      if (fonts.has(entity.instance)) {
        throw new PptParseError("Corrupt legacy PowerPoint file: duplicate font index", {
          part: "record",
          recordType: entity.type,
          recordName: "FontEntityAtom",
          offset: entity.offset,
          reason: "duplicate-font-index",
        });
      }
      fonts.set(entity.instance, name);
    }
  }
  return fonts;
}

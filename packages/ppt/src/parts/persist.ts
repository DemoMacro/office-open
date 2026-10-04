import { PptParseError } from "../errors";
import {
  RecordType,
  collectDescendants,
  readRecordHeader,
  readRecordTree,
  type RecordNode,
} from "../records";

const MAX_USER_EDITS = 256;
const MAX_PERSIST_REFERENCES = 1_000_000;

export function createView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

export function readStreamOffset(
  view: DataView,
  offset: number,
  description: string,
  targetLength = view.byteLength,
): number {
  if (offset < 0 || offset + 4 > view.byteLength) {
    throw new PptParseError(`Corrupt legacy PowerPoint file: invalid ${description}`);
  }
  const value = view.getUint32(offset, true);
  if (value >= targetLength) {
    throw new PptParseError(`Corrupt legacy PowerPoint file: invalid ${description}`);
  }
  return value;
}

export function readPersistReferences(
  view: DataView,
  initialEditOffset: number,
): Map<number, number> {
  const references = new Map<number, number>();
  const visitedEdits = new Set<number>([initialEditOffset]);
  let editOffset: number | undefined = initialEditOffset;

  while (editOffset !== undefined && editOffset !== 0) {
    if (visitedEdits.size > MAX_USER_EDITS) {
      throw new PptParseError("Corrupt legacy PowerPoint file: cyclic UserEditAtom chain");
    }
    const edit = readRecordHeader(view, editOffset);
    if (edit.type !== RecordType.userEditAtom || edit.length < 28) {
      throw new PptParseError("Corrupt legacy PowerPoint file: invalid UserEditAtom chain");
    }
    const pointerOffset = readStreamOffset(view, edit.offset + 20, "persist pointer offset");
    readPersistPointerBlock(view, pointerOffset, references);
    const previousOffset = readStreamOffset(view, edit.offset + 16, "previous UserEditAtom offset");
    if (previousOffset === 0) break;
    if (visitedEdits.has(previousOffset)) {
      throw new PptParseError("Corrupt legacy PowerPoint file: cyclic UserEditAtom chain");
    }
    visitedEdits.add(previousOffset);
    editOffset = previousOffset;
  }

  if (references.size > MAX_PERSIST_REFERENCES) {
    throw new PptParseError("Corrupt legacy PowerPoint file: too many persist objects");
  }
  return references;
}

function readPersistPointerBlock(
  view: DataView,
  offset: number,
  references: Map<number, number>,
): void {
  const block = readRecordHeader(view, offset);
  if (
    block.type !== RecordType.persistPointerIncrementalBlock &&
    block.type !== RecordType.persistPointerFullBlock
  ) {
    throw new PptParseError("Corrupt legacy PowerPoint file: invalid persist pointer block");
  }
  let cursor = block.offset + 8;
  while (cursor + 4 <= block.end) {
    const packed = view.getUint32(cursor, true);
    const firstId = packed & 0xfffff;
    const count = packed >>> 20;
    cursor += 4;
    if (count === 0 || cursor + count * 4 > block.end) {
      throw new PptParseError("Corrupt legacy PowerPoint file: invalid persist pointer table");
    }
    for (let index = 0; index < count; index += 1) {
      const recordOffset = view.getUint32(cursor + index * 4, true);
      references.set(firstId + index, recordOffset);
    }
    cursor += count * 4;
  }
  if (cursor !== block.end) {
    throw new PptParseError("Corrupt legacy PowerPoint file: malformed persist pointer table");
  }
}

export function readPersistedRecord(
  view: DataView,
  offset: number,
  expectedTypes: readonly number[],
): RecordNode {
  const header = readRecordHeader(view, offset);
  const record = readRecordTree(view, offset, header.end)[0];
  if (!record || !expectedTypes.includes(record.type)) {
    throw new PptParseError("Corrupt legacy PowerPoint file: unexpected persist object");
  }
  return record;
}

export function assertNotEncrypted(document: RecordNode): void {
  if (collectDescendants(document, RecordType.documentEncryptionAtom).length > 0) {
    throw new PptParseError("Encrypted legacy PowerPoint files are not supported");
  }
}

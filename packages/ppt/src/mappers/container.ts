import {
  decryptRc4CryptoApi,
  parseRc4CryptoApiHeader,
  verifyRc4CryptoApiPassword,
} from "@office-open/core";

import { readPersistReferences, readStreamOffset } from "../parts/persist";
import {
  LegacyPowerPointError,
  RecordType,
  readRecordHeader,
  readUint32,
  type RecordHeader,
} from "../records";
import { POWERPOINT_DOCUMENT_STREAM } from "./constants";
export function decryptPowerPointDocument(
  encryptedView: DataView,
  userEdit: RecordHeader,
  password?: string,
): Uint8Array {
  if (password === undefined) {
    throw new LegacyPowerPointError("Encrypted legacy PowerPoint files are not supported", {
      part: "stream",
      path: POWERPOINT_DOCUMENT_STREAM,
      recordName: "UserEditAtom",
      offset: userEdit.offset,
      reason: "encrypted-unsupported",
    });
  }
  if (userEdit.length < 32) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: invalid encryption reference",
      {
        part: "record",
        path: POWERPOINT_DOCUMENT_STREAM,
        recordName: "UserEditAtom",
        offset: userEdit.offset,
        length: 32,
        reason: "invalid-record-length",
      },
    );
  }
  const references = readPersistReferences(encryptedView, userEdit.offset);
  const encryptionReference = readUint32(encryptedView, userEdit, 28);
  const encryptionOffset = references.get(encryptionReference);
  if (encryptionOffset === undefined) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: missing encryption record", {
      part: "record",
      path: POWERPOINT_DOCUMENT_STREAM,
      recordType: encryptionReference,
      recordName: "PersistObject",
      reason: "missing-required-record",
    });
  }
  const encryptionRecord = readRecordHeader(encryptedView, encryptionOffset);
  if (encryptionRecord.type !== RecordType.documentEncryptionAtom) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid encryption record", {
      part: "record",
      path: POWERPOINT_DOCUMENT_STREAM,
      recordType: encryptionRecord.type,
      recordName: "PersistObject",
      offset: encryptionRecord.offset,
      length: encryptionRecord.length,
      reason: "unsupported-required-structure",
    });
  }
  const encryptionData = new Uint8Array(
    encryptedView.buffer,
    encryptedView.byteOffset + encryptionRecord.offset + 8,
    encryptionRecord.length,
  );
  const { keySizeBits, verifier } = parseRc4CryptoApiHeader(encryptionData, 4);
  if (!verifyRc4CryptoApiPassword(password, verifier, keySizeBits)) {
    throw new LegacyPowerPointError("Invalid legacy PowerPoint password", {
      part: "record",
      path: POWERPOINT_DOCUMENT_STREAM,
      recordType: RecordType.documentEncryptionAtom,
      recordName: "DocumentEncryptionAtom",
      offset: encryptionRecord.offset,
      reason: "invalid-password",
    });
  }

  const result = new Uint8Array(
    encryptedView.buffer,
    encryptedView.byteOffset,
    encryptedView.byteLength,
  );
  const resultView = new DataView(result.buffer, result.byteOffset, result.byteLength);
  const protectedOffsets = new Set<number>([userEdit.offset]);
  let editOffset: number | undefined = userEdit.offset;
  while (editOffset !== undefined && editOffset !== 0) {
    const edit = readRecordHeader(encryptedView, editOffset);
    protectedOffsets.add(
      readStreamOffset(encryptedView, edit.offset + 20, "persist pointer offset"),
    );
    editOffset = readStreamOffset(encryptedView, edit.offset + 16, "previous UserEditAtom offset");
  }
  const offsets = [...new Set(references.values())].sort((left, right) => left - right);
  const referenceByOffset = new Map(
    [...references].map(([reference, offset]) => [offset, reference]),
  );
  for (const [index, offset] of offsets.entries()) {
    if (protectedOffsets.has(offset)) continue;
    if (offset === encryptionOffset) {
      result.fill(0, offset + 8, offset + 8 + encryptionRecord.length);
      continue;
    }
    const end = index + 1 < offsets.length ? offsets[index + 1]! : encryptedView.byteLength;
    const length = end - offset;
    const recordBlockSize = Math.ceil(length / keySizeBits) * keySizeBits;
    const decrypted = decryptRc4CryptoApi(
      result.subarray(offset, end),
      password,
      verifier.salt,
      keySizeBits,
      recordBlockSize,
      referenceByOffset.get(offset) ?? 0,
    );
    result.set(decrypted, offset);
  }
  const pointerOffset = readStreamOffset(
    encryptedView,
    userEdit.offset + 20,
    "persist pointer offset",
  );
  const pointerBlock = readRecordHeader(encryptedView, pointerOffset);
  const documentReferences = [...references]
    .filter(([reference]) => reference !== encryptionReference)
    .sort(([left], [right]) => left - right);
  const firstReference = documentReferences[0]?.[0] ?? 0;
  const firstOffset = documentReferences[0]?.[1] ?? 0;
  const referenceCount = documentReferences.length;
  const pointerBodyLength = 4 + referenceCount * 4;
  resultView.setUint32(pointerBlock.offset + 4, pointerBodyLength, true);
  resultView.setUint32(pointerBlock.offset + 8, firstReference | (1 << 20), true);
  resultView.setUint32(pointerBlock.offset + 8, firstReference | (referenceCount << 20), true);
  resultView.setUint32(pointerBlock.offset + 12, firstOffset, true);
  for (const [index, [, offset]] of documentReferences.entries()) {
    resultView.setUint32(pointerBlock.offset + 12 + index * 4, offset, true);
  }
  result.fill(0, pointerBlock.offset + 8 + pointerBodyLength, userEdit.offset);
  resultView.setUint32(userEdit.offset + 4, userEdit.length - 4, true);
  resultView.setUint32(userEdit.offset + 8 + 28, 0, true);
  result.fill(0, userEdit.offset + 8 + userEdit.length - 4, userEdit.offset + 8 + userEdit.length);
  return result;
}

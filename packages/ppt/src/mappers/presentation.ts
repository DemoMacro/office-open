import { CompoundFileReader } from "@office-open/core";
import type { PresentationOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";
import { assertCfbSignature, readStream, readSummaryInformation } from "../parts/container";
import {
  assertNotEncrypted,
  createView,
  readPersistedRecord,
  readPersistReferences,
  readStreamOffset,
} from "../parts/persist";
import { RecordType, findDirect, readInt32, readRecordHeader } from "../records";
import {
  CURRENT_USER_STREAM,
  ENCRYPTED_TOKEN,
  NON_ENCRYPTED_TOKEN,
  POWERPOINT_DOCUMENT_STREAM,
} from "./constants";
import { decryptPowerPointDocument } from "./container";
import { readFontNames } from "./fonts";
import type { DrawingContext } from "./models";
import { readPictureStore } from "./pictures";
import { readSlideTextGroups } from "./slide-text";
import { readMaster, readNotesPersistReferences, readSlide } from "./slides";
import { readHyperlinkTargets, masterUnitsToEmu } from "./text";

/**
 * Parse a legacy Microsoft PowerPoint .ppt CFB container into PPTX options.
 * @throws {PptParseError} For non-CFB, corrupt, or encrypted input.
 */
export interface LegacyParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

export function parsePresentation(
  data: Uint8Array,
  options?: LegacyParseOptions,
): PresentationOptions {
  if (!(data instanceof Uint8Array)) {
    throw new PptParseError("parsePresentation expects a Uint8Array", {
      part: "container",
      path: "/",
      reason: "invalid-input-type",
    });
  }
  assertCfbSignature(data);

  let reader: CompoundFileReader;
  try {
    reader = new CompoundFileReader(data);
  } catch (error) {
    throw new PptParseError(
      `Invalid CFB container: ${error instanceof Error ? error.message : String(error)}`,
      { part: "container", path: "/", reason: "invalid-container" },
    );
  }

  let document = readStream(reader, POWERPOINT_DOCUMENT_STREAM);
  let currentUser = readStream(reader, CURRENT_USER_STREAM);
  const pictures = reader.entry("Pictures") ? readPictureStore(reader.read("Pictures")) : [];
  const documentView = createView(document);
  const currentUserView = createView(currentUser);

  if (currentUser.byteLength < 28) {
    throw new PptParseError(
      "Unsupported or corrupt legacy PowerPoint file: Current User atom is too short",
      {
        part: "stream",
        path: CURRENT_USER_STREAM,
        recordType: RecordType.currentUserAtom,
        recordName: "CurrentUserAtom",
        offset: 0,
        length: 28,
        reason: "truncated-header",
      },
    );
  }
  if (currentUserView.getUint16(2, true) !== RecordType.currentUserAtom) {
    throw new PptParseError("Corrupt legacy PowerPoint file: invalid Current User atom", {
      part: "stream",
      path: CURRENT_USER_STREAM,
      recordType: currentUserView.getUint16(2, true),
      recordName: "CurrentUserAtom",
      offset: 2,
      length: 2,
      reason: "unsupported-required-structure",
    });
  }

  const encryptionToken = currentUserView.getUint32(12, true);
  const encrypted = encryptionToken === ENCRYPTED_TOKEN || encryptionToken !== NON_ENCRYPTED_TOKEN;
  if (encrypted) {
    const encryptedEditOffset = readStreamOffset(
      currentUserView,
      16,
      "Current User edit offset",
      documentView.byteLength,
    );
    const encryptedEdit = readRecordHeader(documentView, encryptedEditOffset);
    document = decryptPowerPointDocument(documentView, encryptedEdit, options?.password);
    currentUser = new Uint8Array(currentUser);
    createView(currentUser).setUint32(12, NON_ENCRYPTED_TOKEN, true);
  } else if (encryptionToken !== NON_ENCRYPTED_TOKEN) {
    throw new PptParseError("Encrypted legacy PowerPoint files are not supported", {
      part: "stream",
      path: CURRENT_USER_STREAM,
      recordName: "CurrentUserAtom",
      offset: 12,
      length: 4,
      reason: "encrypted-unsupported",
    });
  }

  const editOffset = readStreamOffset(
    currentUserView,
    16,
    "Current User edit offset",
    documentView.byteLength,
  );
  const userEdit = readRecordHeader(documentView, editOffset);
  if (userEdit.type !== RecordType.userEditAtom || userEdit.length < 28) {
    throw new PptParseError("Corrupt legacy PowerPoint file: invalid UserEditAtom", {
      part: "record",
      path: POWERPOINT_DOCUMENT_STREAM,
      recordType: userEdit.type,
      recordName: "UserEditAtom",
      offset: userEdit.offset,
      length: userEdit.length,
      reason: "unsupported-required-structure",
    });
  }

  const persistReferences = readPersistReferences(documentView, userEdit.offset);
  const documentReference = readInt32(documentView, userEdit, 16);
  const documentOffset = persistReferences.get(documentReference);
  if (documentOffset === undefined) {
    throw new PptParseError("Corrupt legacy PowerPoint file: missing document persist object");
  }

  const documentRecord = readPersistedRecord(documentView, documentOffset, [RecordType.document]);
  assertNotEncrypted(documentRecord);
  const hyperlinkTargets = readHyperlinkTargets(documentView, documentRecord);
  const fontNames = readFontNames(documentView, documentRecord);

  const documentAtom = findDirect(documentRecord, RecordType.documentAtom);
  const size = documentAtom
    ? {
        width: masterUnitsToEmu(readInt32(documentView, documentAtom, 0)),
        height: masterUnitsToEmu(readInt32(documentView, documentAtom, 4)),
      }
    : undefined;

  const slideGroups = readSlideTextGroups(documentView, documentRecord, "slide", fontNames);
  const notesReferences = readNotesPersistReferences(documentView, persistReferences);
  const masterGroups = readSlideTextGroups(documentView, documentRecord, "master", fontNames);
  const masterNames = new Map<number, string>();
  const masters = masterGroups.map((group, index) => {
    const name = `Legacy Master ${index + 1}`;
    if (group.slideIdentifier !== undefined && !masterNames.has(group.slideIdentifier)) {
      masterNames.set(group.slideIdentifier, name);
    }
    return readMaster(
      documentView,
      persistReferences,
      group,
      {
        entries: group.entries,
        pictures,
        usedEntryIndexes: new Set(),
        hyperlinks: hyperlinkTargets,
        fonts: fontNames,
      },
      name,
    );
  });
  const defaultMaster = masters[0]?.name;
  const slides = slideGroups.map((group) => {
    const context: DrawingContext = {
      entries: group.entries,
      pictures,
      usedEntryIndexes: new Set(),
      hyperlinks: hyperlinkTargets,
      fonts: fontNames,
    };
    return readSlide(
      documentView,
      persistReferences,
      group,
      context,
      defaultMaster,
      masterNames,
      notesReferences,
    );
  });
  const result: PresentationOptions = {};
  Object.assign(result, readSummaryInformation(reader));
  if (size) result.size = size;
  if (masters.length > 0) result.masters = masters;
  if (slides.length > 0) result.slides = slides;
  return result;
}

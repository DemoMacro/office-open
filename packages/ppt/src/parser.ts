import { CompoundFileReader } from "@office-open/core";
import type { TextBodyOptions, TextRunOptions } from "@office-open/core";
import type {
  PresentationOptions,
  ShapeOptions,
  SlideChild,
  SlideOptions,
} from "@office-open/pptx";

import {
  LegacyPowerPointError,
  RecordType,
  collectDescendants,
  decodeAnsi,
  decodeUtf16,
  findDescendant,
  findDirect,
  readInt16,
  readInt32,
  readRecordHeader,
  readRecordTree,
  recordBody,
  type RecordNode,
} from "./records";

const POWERPOINT_DOCUMENT_STREAM = "PowerPoint Document";
const CURRENT_USER_STREAM = "Current User";
const NON_ENCRYPTED_TOKEN = 0xe391c05f;
const ENCRYPTED_TOKEN = 0xf3d1c4df;
const EMU_PER_INCH = 914400;
const MASTER_UNITS_PER_INCH = 576;
const MAX_USER_EDITS = 256;
const MAX_PERSIST_REFERENCES = 1_000_000;

interface TextEntry {
  readonly placeholderType: number;
  readonly text: string;
}

interface SlideTextGroup {
  readonly persistReference: number;
  readonly entries: TextEntry[];
}

/**
 * Parse a legacy Microsoft PowerPoint .ppt CFB container into PPTX options.
 * @throws {LegacyPowerPointError} For non-CFB, corrupt, or encrypted input.
 */
export function parsePresentation(data: Uint8Array): PresentationOptions {
  if (!(data instanceof Uint8Array)) throw new TypeError("parsePresentation expects a Uint8Array");
  assertCfbSignature(data);

  let reader: CompoundFileReader;
  try {
    reader = new CompoundFileReader(data);
  } catch (error) {
    throw new LegacyPowerPointError(
      `Invalid CFB container: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const document = readStream(reader, POWERPOINT_DOCUMENT_STREAM);
  const currentUser = readStream(reader, CURRENT_USER_STREAM);
  const documentView = createView(document);
  const currentUserView = createView(currentUser);

  if (currentUser.byteLength < 28) {
    throw new LegacyPowerPointError(
      "Unsupported or corrupt legacy PowerPoint file: Current User atom is too short",
    );
  }
  if (currentUserView.getUint16(2, true) !== RecordType.currentUserAtom) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid Current User atom");
  }

  const encryptionToken = currentUserView.getUint32(12, true);
  if (encryptionToken === ENCRYPTED_TOKEN || encryptionToken !== NON_ENCRYPTED_TOKEN) {
    throw new LegacyPowerPointError("Encrypted legacy PowerPoint files are not supported");
  }

  const editOffset = readStreamOffset(
    currentUserView,
    16,
    "Current User edit offset",
    documentView.byteLength,
  );
  const userEdit = readRecordHeader(documentView, editOffset);
  if (userEdit.type !== RecordType.userEditAtom || userEdit.length < 28) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid UserEditAtom");
  }

  const persistReferences = readPersistReferences(documentView, userEdit.offset);
  const documentReference = readInt32(documentView, userEdit, 16);
  const documentOffset = persistReferences.get(documentReference);
  if (documentOffset === undefined) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: missing document persist object",
    );
  }

  const documentRecord = readPersistedRecord(documentView, documentOffset, [RecordType.document]);
  assertNotEncrypted(documentRecord);

  const documentAtom = findDirect(documentRecord, RecordType.documentAtom);
  const size = documentAtom
    ? {
        width: masterUnitsToEmu(readInt32(documentView, documentAtom, 0)),
        height: masterUnitsToEmu(readInt32(documentView, documentAtom, 4)),
      }
    : undefined;

  const slideGroups = readSlideTextGroups(documentView, documentRecord);
  const slides = slideGroups.map((group) => readSlide(documentView, persistReferences, group));
  const result: PresentationOptions = {};
  if (size) result.size = size;
  if (slides.length > 0) result.slides = slides;
  return result;
}

function assertCfbSignature(data: Uint8Array): void {
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (data.byteLength < signature.length || signature.some((byte, index) => data[index] !== byte)) {
    throw new LegacyPowerPointError("Input is not a legacy PowerPoint CFB container");
  }
}

function readStream(reader: CompoundFileReader, path: string): Uint8Array {
  try {
    return reader.read(path);
  } catch {
    throw new LegacyPowerPointError(
      `Invalid legacy PowerPoint CFB container: missing or unreadable "${path}" stream`,
    );
  }
}

function createView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function readStreamOffset(
  view: DataView,
  offset: number,
  description: string,
  targetLength = view.byteLength,
): number {
  if (offset < 0 || offset + 4 > view.byteLength) {
    throw new LegacyPowerPointError(`Corrupt legacy PowerPoint file: invalid ${description}`);
  }
  const value = view.getUint32(offset, true);
  if (value >= targetLength) {
    throw new LegacyPowerPointError(`Corrupt legacy PowerPoint file: invalid ${description}`);
  }
  return value;
}

function readPersistReferences(view: DataView, initialEditOffset: number): Map<number, number> {
  const references = new Map<number, number>();
  const visitedEdits = new Set<number>([initialEditOffset]);
  let editOffset: number | undefined = initialEditOffset;
  let visitedPointers = 0;

  while (editOffset !== undefined && editOffset !== 0) {
    if (visitedEdits.size > MAX_USER_EDITS) {
      throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: cyclic UserEditAtom chain");
    }
    const edit = readRecordHeader(view, editOffset);
    if (edit.type !== RecordType.userEditAtom || edit.length < 28) {
      throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid UserEditAtom chain");
    }
    const pointerOffset = readStreamOffset(view, edit.offset + 20, "persist pointer offset");
    readPersistPointerBlock(view, pointerOffset, references);
    visitedPointers += 1;
    const previousOffset = readStreamOffset(view, edit.offset + 16, "previous UserEditAtom offset");
    if (previousOffset === 0) break;
    if (visitedEdits.has(previousOffset)) {
      throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: cyclic UserEditAtom chain");
    }
    visitedEdits.add(previousOffset);
    editOffset = previousOffset;
  }

  if (references.size > MAX_PERSIST_REFERENCES) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: too many persist objects");
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
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: invalid persist pointer block",
    );
  }
  let cursor = block.offset + 8;
  while (cursor + 4 <= block.end) {
    const packed = view.getUint32(cursor, true);
    const firstId = packed & 0xfffff;
    const count = packed >>> 20;
    cursor += 4;
    if (count === 0 || cursor + count * 4 > block.end) {
      throw new LegacyPowerPointError(
        "Corrupt legacy PowerPoint file: invalid persist pointer table",
      );
    }
    for (let index = 0; index < count; index += 1) {
      const recordOffset = view.getUint32(cursor + index * 4, true);
      const record = readRecordHeader(view, recordOffset);
      references.set(firstId + index, record.offset);
    }
    cursor += count * 4;
  }
  if (cursor !== block.end) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: malformed persist pointer table",
    );
  }
}

function readPersistedRecord(
  view: DataView,
  offset: number,
  expectedTypes: readonly number[],
): RecordNode {
  const record = readRecordTree(view, offset, view.byteLength)[0];
  if (!record || !expectedTypes.includes(record.type)) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: unexpected persist object");
  }
  return record;
}

function assertNotEncrypted(document: RecordNode): void {
  if (collectDescendants(document, RecordType.documentEncryptionAtom).length > 0) {
    throw new LegacyPowerPointError("Encrypted legacy PowerPoint files are not supported");
  }
}

function readSlideTextGroups(view: DataView, document: RecordNode): SlideTextGroup[] {
  const lists = document.children
    .filter((child) => child.type === RecordType.slideListWithText && child.instance === 0)
    .flatMap((list) => groupSlideText(view, list));
  return lists;
}

function groupSlideText(view: DataView, list: RecordNode): SlideTextGroup[] {
  const groups: SlideTextGroup[] = [];
  let current: { persistReference: number; entries: TextEntry[] } | undefined;
  let placeholderType: number | undefined;

  for (const child of list.children) {
    if (child.type === RecordType.slidePersistAtom) {
      if (current) groups.push(current);
      current = { persistReference: readInt32(view, child, 0), entries: [] };
      placeholderType = undefined;
      continue;
    }
    if (child.type === RecordType.textHeader) {
      placeholderType = readInt32(view, child, 0);
      continue;
    }
    if (current && placeholderType !== undefined) {
      const bytes =
        child.type === RecordType.textChars || child.type === RecordType.textBytes
          ? recordBody(view, child)
          : undefined;
      if (bytes) {
        current.entries.push({
          placeholderType,
          text: child.type === RecordType.textChars ? decodeUtf16(bytes) : decodeAnsi(bytes),
        });
        placeholderType = undefined;
      }
    }
  }
  if (current) groups.push(current);
  return groups;
}

function readSlide(
  view: DataView,
  references: Map<number, number>,
  group: SlideTextGroup,
): SlideOptions {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: missing slide persist object");
  }
  const slide = readRecordTree(view, offset, view.byteLength)[0];
  if (!slide || slide.type !== RecordType.slide) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected slide persist object",
    );
  }
  const drawing = findDirect(slide, RecordType.ppDrawing);
  const children = drawing ? readSlideShapes(view, drawing, group.entries) : [];
  const result: SlideOptions = {};
  if (children.length > 0) result.children = children;
  return result;
}

function readSlideShapes(view: DataView, drawing: RecordNode, entries: TextEntry[]): SlideChild[] {
  const shapeContainers = collectDescendants(drawing, RecordType.escherShapeContainer);
  const children: SlideChild[] = [];
  for (const container of shapeContainers) {
    const shape = parseShape(view, container, entries);
    if (shape) children.push({ shape });
  }
  return children;
}

function parseShape(
  view: DataView,
  container: RecordNode,
  entries: TextEntry[],
): ShapeOptions | undefined {
  const shapeRecord = findDirect(container, RecordType.escherShape);
  if (!shapeRecord) return undefined;
  const anchor = findDirect(container, RecordType.escherClientAnchor);
  if (!anchor || anchor.length < 8) return undefined;
  const text = readContainerText(view, container, entries);
  if (!text) return undefined;

  const shape: ShapeOptions = {
    ...(shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(readInt16(view, anchor, 2)),
    y: masterUnitsToEmu(readInt16(view, anchor, 0)),
    width: masterUnitsToEmu(readInt16(view, anchor, 4) - readInt16(view, anchor, 2)),
    height: masterUnitsToEmu(readInt16(view, anchor, 6) - readInt16(view, anchor, 0)),
    textBox: true,
    textBody: text,
  };
  return shape;
}

function readContainerText(
  view: DataView,
  container: RecordNode,
  slideEntries: TextEntry[],
): TextBodyOptions | undefined {
  const clientTextbox = findDirect(container, RecordType.escherClientTextbox);
  if (!clientTextbox) return undefined;

  const embeddedText = collectEmbeddedText(view, clientTextbox);
  if (embeddedText.length > 0) return createTextBody(embeddedText);

  const reference = findDescendant(clientTextbox, RecordType.outlineTextReference);
  if (reference && reference.length >= 4) {
    const index = readInt32(view, reference, 0);
    const entry = slideEntries[index];
    if (entry) return createTextBody([entry.text]);
  }
  return undefined;
}

function collectEmbeddedText(view: DataView, parent: RecordNode): string[] {
  const text: string[] = [];
  let placeholderType: number | undefined;
  for (const child of flatten(parent)) {
    if (child.type === RecordType.textHeader) {
      placeholderType = readInt32(view, child, 0);
    } else if (placeholderType !== undefined && child.type === RecordType.textChars) {
      text.push(decodeUtf16(recordBody(view, child)));
      placeholderType = undefined;
    } else if (placeholderType !== undefined && child.type === RecordType.textBytes) {
      text.push(decodeAnsi(recordBody(view, child)));
      placeholderType = undefined;
    }
  }
  return text;
}

function flatten(record: RecordNode): RecordNode[] {
  const output: RecordNode[] = [];
  const visit = (node: RecordNode): void => {
    for (const child of node.children) {
      output.push(child);
      visit(child);
    }
  };
  visit(record);
  return output;
}

function createTextBody(text: readonly string[]): TextBodyOptions {
  const paragraphs = text
    .flatMap((value) => splitParagraphs(value))
    .map((value) => ({ children: [{ text: value }] satisfies TextRunOptions[] }));
  return paragraphs.length > 0 ? { paragraphs } : {};
}

function splitParagraphs(value: string): string[] {
  return value.split(/\r\n|\r|\n/);
}

function masterUnitsToEmu(value: number): number {
  return Math.round((value * EMU_PER_INCH) / MASTER_UNITS_PER_INCH);
}

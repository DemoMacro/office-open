import {
  CompoundFileReader,
  decryptRc4CryptoApi,
  parseRc4CryptoApiHeader,
  parseSummaryInformation,
  verifyRc4CryptoApiPassword,
} from "@office-open/core";
import type { TextBodyOptions, TextRunOptions } from "@office-open/core";
import type {
  AnimationEntry,
  GroupOptions,
  MasterDefinition,
  PresentationOptions,
  PictureOptions,
  ShapeOptions,
  SlideChild,
  SlideOptions,
  TableOptions,
} from "@office-open/pptx";

import {
  LegacyPowerPointError,
  RecordType,
  collectDescendants,
  decodeAnsi,
  decodeUtf16,
  findDescendant,
  findDirect,
  isZeroPadding,
  readInt16,
  readInt32,
  readUint32,
  readRecordHeader,
  readRecordTree,
  recordBody,
  type RecordNode,
  type RecordHeader,
} from "./records";

const POWERPOINT_DOCUMENT_STREAM = "PowerPoint Document";
const CURRENT_USER_STREAM = "Current User";
const NON_ENCRYPTED_TOKEN = 0xe391c05f;
const ENCRYPTED_TOKEN = 0xf3d1c4df;
const EMU_PER_INCH = 914400;
const MASTER_UNITS_PER_INCH = 576;
const MAX_USER_EDITS = 256;
const MAX_PERSIST_REFERENCES = 1_000_000;
const TABLE_BOUND_TOLERANCE = 2;
const RECORD_BODY_OFFSET = 8;
const PICTURE_RECORD_TYPES = new Set([
  0xf01a, 0xf01b, 0xf01c, 0xf01d, 0xf01e, 0xf01f, 0xf029, 0xf02a,
]);

interface TextEntry {
  readonly placeholderType: number;
  readonly text: string;
}

type TableCell = NonNullable<TableOptions["rows"]>[number]["cells"][number];

interface SlideTextGroup {
  readonly persistReference: number;
  readonly slideIdentifier: number | undefined;
  readonly entries: TextEntry[];
}

interface LegacyPicture {
  readonly type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  readonly data: Uint8Array;
}

interface DrawingContext {
  readonly entries: TextEntry[];
  readonly pictures: readonly (LegacyPicture | undefined)[];
  readonly usedEntryIndexes: Set<number>;
  readonly hyperlinks: ReadonlyMap<number, { url?: string; tooltip?: string }>;
}

interface LegacyTextHyperlink {
  readonly start: number;
  readonly end: number;
  readonly target: { url?: string; tooltip?: string };
}

interface LegacyAnimation {
  readonly shapeId: number | undefined;
  readonly atom: RecordNode;
}

/**
 * Parse a legacy Microsoft PowerPoint .ppt CFB container into PPTX options.
 * @throws {LegacyPowerPointError} For non-CFB, corrupt, or encrypted input.
 */
export interface LegacyParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

export function parsePresentation(
  data: Uint8Array,
  options?: LegacyParseOptions,
): PresentationOptions {
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

  let document = readStream(reader, POWERPOINT_DOCUMENT_STREAM);
  let currentUser = readStream(reader, CURRENT_USER_STREAM);
  const pictures = reader.entry("Pictures") ? readPictureStore(reader.read("Pictures")) : [];
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
  const hyperlinkTargets = readHyperlinkTargets(documentView, documentRecord);

  const documentAtom = findDirect(documentRecord, RecordType.documentAtom);
  const size = documentAtom
    ? {
        width: masterUnitsToEmu(readInt32(documentView, documentAtom, 0)),
        height: masterUnitsToEmu(readInt32(documentView, documentAtom, 4)),
      }
    : undefined;

  const slideGroups = readSlideTextGroups(documentView, documentRecord, "slide");
  const notesGroups = readSlideTextGroups(documentView, documentRecord, "notes");
  const masterGroups = readSlideTextGroups(documentView, documentRecord, "master");
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
    };
    return readSlide(
      documentView,
      persistReferences,
      group,
      context,
      defaultMaster,
      masterNames,
      notesGroups,
    );
  });
  const result: PresentationOptions = {};
  Object.assign(result, readSummaryInformation(reader));
  if (size) result.size = size;
  if (masters.length > 0) result.masters = masters;
  if (slides.length > 0) result.slides = slides;
  return result;
}

function decryptPowerPointDocument(
  encryptedView: DataView,
  userEdit: RecordHeader,
  password?: string,
): Uint8Array {
  if (password === undefined) {
    throw new LegacyPowerPointError("Encrypted legacy PowerPoint files are not supported");
  }
  if (userEdit.length < 32) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid encryption reference");
  }
  const references = readPersistReferences(encryptedView, userEdit.offset);
  const encryptionReference = readUint32(encryptedView, userEdit, 28);
  const encryptionOffset = references.get(encryptionReference);
  if (encryptionOffset === undefined) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: missing encryption record");
  }
  const encryptionRecord = readRecordHeader(encryptedView, encryptionOffset);
  if (encryptionRecord.type !== RecordType.documentEncryptionAtom) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: invalid encryption record");
  }
  const encryptionData = new Uint8Array(
    encryptedView.buffer,
    encryptedView.byteOffset + encryptionRecord.offset + 8,
    encryptionRecord.length,
  );
  const { keySizeBits, verifier } = parseRc4CryptoApiHeader(encryptionData, 4);
  if (!verifyRc4CryptoApiPassword(password, verifier, keySizeBits)) {
    throw new LegacyPowerPointError("Invalid legacy PowerPoint password");
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

function readSummaryInformation(reader: CompoundFileReader): PresentationOptions {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  return reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
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
      references.set(firstId + index, recordOffset);
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
  const header = readRecordHeader(view, offset);
  const record = readRecordTree(view, offset, header.end)[0];
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

function readSlideTextGroups(
  view: DataView,
  document: RecordNode,
  kind: "slide" | "master" | "notes",
): SlideTextGroup[] {
  const lists = document.children
    .filter(
      (child) =>
        child.type === RecordType.slideListWithText &&
        child.instance === (kind === "slide" ? 0 : kind === "master" ? 1 : 2),
    )
    .flatMap((list) => groupSlideText(view, list));
  return lists;
}

function groupSlideText(view: DataView, list: RecordNode): SlideTextGroup[] {
  const groups: SlideTextGroup[] = [];
  let current: SlideTextGroup | undefined;
  let placeholderType: number | undefined;

  for (const child of list.children) {
    if (child.type === RecordType.slidePersistAtom) {
      if (current) groups.push(current);
      current = {
        persistReference: readInt32(view, child, 0),
        slideIdentifier: child.length >= 16 ? readUint32(view, child, 12) : undefined,
        entries: [],
      };
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
  context: DrawingContext,
  defaultMaster: string | undefined,
  masterNames: ReadonlyMap<number, string>,
  notesGroups: readonly SlideTextGroup[],
): SlideOptions {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: missing slide persist object");
  }
  const slideHeader = readRecordHeader(view, offset);
  const slide = readRecordTree(view, offset, slideHeader.end)[0];
  if (!slide || slide.type !== RecordType.slide) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected slide persist object",
    );
  }
  const drawing = findDirect(slide, RecordType.ppDrawing);
  const drawingChildren = drawing ? readDrawingChildren(view, drawing, context) : [];
  const children = appendUnreferencedText(drawingChildren, context);
  const slideAtom = findDirect(slide, RecordType.slideAtom);
  const hasSlideAtom = slideAtom !== undefined && slideAtom.length >= 24;
  const masterId = hasSlideAtom ? readUint32(view, slideAtom!, 16) : undefined;
  const notesId = hasSlideAtom ? readUint32(view, slideAtom!, 20) : undefined;
  const master = masterId !== undefined && masterId > 0 ? masterNames.get(masterId) : undefined;
  const result: SlideOptions = {};
  if (children.length > 0) result.children = children;
  const slideMaster = master ?? defaultMaster;
  if (slideMaster) result.master = slideMaster;
  const notes = readNotes(view, references, notesGroups, notesId);
  if (notes) result.notes = notes;
  const animations = drawing
    ? readSlideAnimations(view, drawing).map((animation) => projectAnimation(view, animation))
    : [];
  if (animations.length > 0) result.animations = animations;
  return result;
}

function readNotes(
  view: DataView,
  references: Map<number, number>,
  notesGroups: readonly SlideTextGroup[],
  notesId: number | undefined,
): string | undefined {
  if (!notesId) return undefined;
  const group = notesGroups.find((entry) => entry.slideIdentifier === notesId);
  const offset = group && references.get(group.persistReference);
  if (offset === undefined) return undefined;
  try {
    const notes = readPersistedRecord(view, offset, [RecordType.notes]);
    const drawing = findDirect(notes, RecordType.ppDrawing);
    if (!drawing) return undefined;
    const paragraphs = collectEmbeddedText(view, drawing)
      .flatMap((value) => splitParagraphs(value))
      .filter((value) => value.length > 0);
    return paragraphs.length > 0 ? paragraphs.join("\n") : undefined;
  } catch {
    return undefined;
  }
}

function readSlideAnimations(view: DataView, drawing: RecordNode): LegacyAnimation[] {
  const shapes = collectDescendants(drawing, RecordType.escherShapeContainer);
  const animations: LegacyAnimation[] = [];
  for (const shape of shapes) {
    try {
      const shapeRecord = findDirect(shape, RecordType.escherShape);
      const container = findDirect(shape, RecordType.animationInfo);
      const atom = container && findDirect(container, RecordType.animationInfoAtom);
      if (atom && atom.length >= 28) {
        animations.push({
          shapeId:
            shapeRecord && shapeRecord.length >= 4 ? readInt32(view, shapeRecord, 0) : undefined,
          atom,
        });
      }
    } catch {
      continue;
    }
  }
  return animations.sort(
    (first, second) => animationOrder(view, first) - animationOrder(view, second),
  );
}

function animationOrder(view: DataView, animation: LegacyAnimation): number {
  try {
    return readInt16(view, animation.atom, 16);
  } catch {
    return 0;
  }
}

function projectAnimation(view: DataView, animation: LegacyAnimation): AnimationEntry {
  const atom = animation.atom;
  const effect = view.getUint8(atom.offset + RECORD_BODY_OFFSET + 21);
  const mask = view.getUint16(atom.offset + RECORD_BODY_OFFSET + 4, true);
  const automatic = ((mask >>> 2) & 1) === 1;
  let delay = 0;
  try {
    delay = Math.max(0, readInt32(view, atom, 12));
  } catch {
    delay = 0;
  }
  return {
    type: animationType(effect),
    class: "entrance",
    trigger: automatic ? "afterPrevious" : "onClick",
    ...(delay > 0 ? { delay } : {}),
    ...(animation.shapeId !== undefined && animation.shapeId > 0
      ? { shapeId: animation.shapeId }
      : {}),
  };
}

function animationType(effect: number): AnimationEntry["type"] {
  if (effect === 0x02) return "blinds";
  if (effect === 0x03) return "checker";
  if (effect === 0x04) return "cover";
  if (effect === 0x05) return "dissolve";
  if (effect === 0x06) return "fade";
  if (effect === 0x08) return "randomBars";
  if (effect === 0x09) return "strips";
  if (effect === 0x0a) return "wipe";
  if (effect === 0x0b) return "fly";
  if (effect === 0x0d) return "split";
  if (effect === 0x1a) return "wheel";
  return "appear";
}

function readMaster(
  view: DataView,
  references: Map<number, number>,
  group: SlideTextGroup,
  context: DrawingContext,
  name: string,
): MasterDefinition {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: missing master persist object",
    );
  }
  const masterHeader = readRecordHeader(view, offset);
  const master = readRecordTree(view, offset, masterHeader.end)[0];
  if (!master || (master.type !== RecordType.slide && master.type !== RecordType.mainMaster)) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected master persist object",
    );
  }
  const drawing = findDirect(master, RecordType.ppDrawing);
  const drawingChildren = drawing ? readDrawingChildren(view, drawing, context) : [];
  const children = appendUnreferencedText(drawingChildren, context);
  return children.length > 0 ? { name, children } : { name };
}

function appendUnreferencedText(children: SlideChild[], context: DrawingContext): SlideChild[] {
  const result = [...children];
  const projectedText = children.flatMap(slideChildText).map(compactText);
  for (const [index, entry] of context.entries.entries()) {
    if (context.usedEntryIndexes.has(index)) continue;
    if (projectedText.includes(compactText(entry.text))) continue;
    result.push({
      shape: {
        textBox: true,
        textBody: createTextBody([entry.text]),
      },
    });
  }
  return result;
}

function slideChildText(child: SlideChild): string[] {
  if ("shape" in child) return [textBodyToText(child.shape.textBody)];
  if ("group" in child) return child.group.children.flatMap(slideChildText);
  if ("table" in child) {
    return child.table.rows.flatMap((row) =>
      row.cells.map((cell) => (typeof cell.text === "string" ? cell.text : textBodyToText(cell))),
    );
  }
  return [];
}

function compactText(value: string): string {
  return value.replace(/\s+/g, "");
}

function readPictureStore(data: Uint8Array): readonly (LegacyPicture | undefined)[] {
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

function readDrawingChildren(
  view: DataView,
  drawing: RecordNode,
  context: DrawingContext,
): SlideChild[] {
  const pageContainer = findDirect(drawing, RecordType.escherDrawingContainer);
  const shapeGroups = pageContainer
    ? pageContainer.children.filter((child) => child.type === RecordType.escherShapeGroupContainer)
    : [];
  const children: SlideChild[] = [];
  for (const shapeGroup of shapeGroups) {
    children.push(...parseShapeGroup(view, shapeGroup, context, 0));
  }
  return children;
}

function parseShapeGroup(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
  depth: number,
): SlideChild[] {
  if (depth > 16) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint drawing: group depth limit exceeded",
    );
  }

  const children: SlideChild[] = [];
  const childRecords = container.children;
  const metadata = childRecords.find(
    (child) =>
      child.type === RecordType.escherShapeContainer &&
      findDirect(child, RecordType.escherShapeGroup) !== undefined,
  );

  for (const child of childRecords) {
    if (child.type === RecordType.escherShapeGroupContainer) {
      children.push(...parseShapeGroup(view, child, context, depth + 1));
      continue;
    }
    if (child.type !== RecordType.escherShapeContainer || child === metadata) continue;
    const shape = parseShape(view, child, context);
    if (!shape) continue;
    if ("data" in shape) children.push({ picture: shape });
    else children.push({ shape });
  }

  if (children.length === 0) return [];
  if (depth === 0 && children.length === 1) {
    const only = children[0]!;
    if ("table" in only) return children;
    if ("group" in only) {
      const nested = only.group.children;
      if (nested.length === 1 && "table" in nested[0]!) return nested;
    }
  }

  const table = detectGroupedTable(view, childRecords, context);
  if (table) return [{ table }];

  const group = parseGroupShape(view, metadata, children);
  return group ? [{ group }] : children;
}

function detectGroupedTable(
  view: DataView,
  records: readonly RecordNode[],
  context: DrawingContext,
): TableOptions | undefined {
  const metadata = records.find(
    (record) =>
      record.type === RecordType.escherShapeContainer &&
      findDirect(record, RecordType.escherShapeGroup) !== undefined,
  );
  const frame = readTableFrame(view, metadata);
  const groupBounds = readGroupBounds(view, metadata);
  const shapeRecords = records.filter(
    (record) => record.type === RecordType.escherShapeContainer && record !== metadata,
  );
  if (shapeRecords.some((record) => !findDirect(record, RecordType.escherChildAnchor))) {
    return undefined;
  }

  const bounds = shapeRecords.flatMap(
    (
      record,
    ): {
      shape: ShapeOptions;
      left: number;
      top: number;
      right: number;
      bottom: number;
    }[] => {
      const shape = parseShape(view, record, context);
      const anchor = readAnchorBounds(view, record);
      if (!shape || !("textBody" in shape) || !anchor) return [];
      return [{ shape, ...anchor }];
    },
  );
  if (!frame || frame.width <= 0 || frame.height <= 0) return undefined;
  if (bounds.some((item) => item.left >= item.right || item.top >= item.bottom)) return undefined;
  if (groupBounds) {
    const childLeft = Math.min(groupBounds.childLeft, groupBounds.childRight);
    const childTop = Math.min(groupBounds.childTop, groupBounds.childBottom);
    const childRight = Math.max(groupBounds.childLeft, groupBounds.childRight);
    const childBottom = Math.max(groupBounds.childTop, groupBounds.childBottom);
    if (
      bounds.some(
        (item) =>
          item.left < childLeft - TABLE_BOUND_TOLERANCE ||
          item.top < childTop - TABLE_BOUND_TOLERANCE ||
          item.right > childRight + TABLE_BOUND_TOLERANCE ||
          item.bottom > childBottom + TABLE_BOUND_TOLERANCE,
      )
    ) {
      return undefined;
    }
  }

  const lefts = uniqueSorted(bounds.map((item) => item.left));
  const tops = uniqueSorted(bounds.map((item) => item.top));
  const rightEdges = uniqueSorted(bounds.map((item) => item.right));
  const bottomEdges = uniqueSorted(bounds.map((item) => item.bottom));
  const columnEdges = uniqueSorted([...lefts, ...rightEdges]);
  const rowEdges = uniqueSorted([...tops, ...bottomEdges]);
  if (
    columnEdges.length < 3 ||
    tops.length < 2 ||
    columnEdges.length !== lefts.length + 1 ||
    rowEdges.length !== tops.length + 1 ||
    lefts.some((left, index) => !near(left, columnEdges[index]!)) ||
    rightEdges.some((right, index) => !near(right, columnEdges[index + 1]!)) ||
    tops.some((top, index) => !near(top, rowEdges[index]!)) ||
    bottomEdges.some((bottom, index) => !near(bottom, rowEdges[index + 1]!))
  ) {
    return undefined;
  }

  const columnCount = columnEdges.length - 1;
  const rowCount = rowEdges.length - 1;
  const occupied = Array.from({ length: rowCount }, () => Array<boolean>(columnCount).fill(false));
  const cells: (TableCell | undefined)[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: columnCount }, () => undefined),
  );

  for (const item of bounds.sort(
    (first, second) => first.top - second.top || first.left - second.left,
  )) {
    const columnStart = columnEdges.findIndex((edge) => near(edge, item.left));
    const columnEnd = columnEdges.findIndex((edge) => near(edge, item.right));
    const rowStart = rowEdges.findIndex((edge) => near(edge, item.top));
    const rowEnd = rowEdges.findIndex((edge) => near(edge, item.bottom));
    const columnSpan = columnEnd - columnStart;
    const rowSpan = rowEnd - rowStart;
    if (columnStart < 0 || rowStart < 0 || columnSpan <= 0 || rowSpan <= 0) return undefined;

    let firstColumn = -1;
    let firstRow = -1;
    outer: for (let row = rowStart; row < rowEnd; row += 1) {
      for (let column = columnStart; column < columnEnd; column += 1) {
        if (!occupied[row]![column]) {
          firstRow = row;
          firstColumn = column;
          break outer;
        }
      }
    }
    if (firstRow < 0 || firstColumn < 0) return undefined;
    for (let row = firstRow; row < firstRow + rowSpan; row += 1) {
      for (let column = firstColumn; column < firstColumn + columnSpan; column += 1) {
        if (occupied[row]![column]) return undefined;
        occupied[row]![column] = true;
      }
    }
    cells[firstRow]![firstColumn] = {
      ...tableCell(item.shape),
      ...(columnSpan > 1 ? { columnSpan } : {}),
      ...(rowSpan > 1 ? { rowSpan } : {}),
    };
  }
  if (occupied.some((row) => row.some((cell) => !cell))) return undefined;

  const columnWidths = columnEdges
    .slice(0, -1)
    .map((edge, index) => masterUnitsToEmu(columnEdges[index + 1]! - edge));
  const rowHeights = rowEdges
    .slice(0, -1)
    .map((edge, index) => masterUnitsToEmu(rowEdges[index + 1]! - edge));

  const rows: NonNullable<TableOptions["rows"]> = [];
  for (const [index, height] of rowHeights.entries()) {
    if (height <= 0) return undefined;
    rows.push({ height, cells: cells[index]!.filter((cell) => cell !== undefined) });
  }

  return { ...frame, rows, columnWidths };
}

function readAnchorBounds(
  view: DataView,
  container: RecordNode,
): { left: number; top: number; right: number; bottom: number } | undefined {
  const clientAnchor = findDirect(container, RecordType.escherClientAnchor);
  if (clientAnchor) {
    if (clientAnchor.length < 8) return undefined;
    return {
      top: readInt16(view, clientAnchor, 0),
      left: readInt16(view, clientAnchor, 2),
      right: readInt16(view, clientAnchor, 4),
      bottom: readInt16(view, clientAnchor, 6),
    };
  }

  const childAnchor = findDirect(container, RecordType.escherChildAnchor);
  if (!childAnchor || childAnchor.length < 16) return undefined;
  return {
    left: readInt32(view, childAnchor, 0),
    top: readInt32(view, childAnchor, 4),
    right: readInt32(view, childAnchor, 8),
    bottom: readInt32(view, childAnchor, 12),
  };
}

function uniqueSorted(values: readonly number[]): number[] {
  return [...new Set(values)].sort((first, second) => first - second);
}

function near(first: number, second: number): boolean {
  return Math.abs(first - second) <= TABLE_BOUND_TOLERANCE;
}

function textBodyToText(body: TextBodyOptions | undefined): string {
  if (body?.text !== undefined) return body.text;
  return (body?.paragraphs ?? [])
    .map((paragraph) =>
      typeof paragraph === "string"
        ? paragraph
        : (paragraph.children ?? [])
            .map((child) =>
              typeof child === "string" ? child : "text" in child ? (child.text ?? "") : "",
            )
            .join(""),
    )
    .join("\n");
}

function parseGroupShape(
  view: DataView,
  metadata: RecordNode | undefined,
  children: SlideChild[],
): GroupOptions | undefined {
  const transform = readGroupTransform(view, metadata);
  if (!transform) return undefined;
  return { ...transform, children };
}

interface GroupBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly childLeft: number;
  readonly childTop: number;
  readonly childRight: number;
  readonly childBottom: number;
}

interface TableFrame {
  readonly id?: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

function readGroupTransform(
  view: DataView,
  metadata: RecordNode | undefined,
):
  | Pick<
      GroupOptions,
      | "id"
      | "x"
      | "y"
      | "width"
      | "height"
      | "childOffsetX"
      | "childOffsetY"
      | "childExtentWidth"
      | "childExtentHeight"
    >
  | undefined {
  const bounds = readGroupBounds(view, metadata);
  if (!bounds) return undefined;
  const shapeRecord = metadata && findDirect(metadata, RecordType.escherShape);
  return {
    ...(shapeRecord && shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
    childOffsetX: masterUnitsToEmu(Math.min(bounds.childLeft, bounds.childRight)),
    childOffsetY: masterUnitsToEmu(Math.min(bounds.childTop, bounds.childBottom)),
    childExtentWidth: masterUnitsToEmu(Math.abs(bounds.childRight - bounds.childLeft)),
    childExtentHeight: masterUnitsToEmu(Math.abs(bounds.childBottom - bounds.childTop)),
  };
}

function readGroupBounds(
  view: DataView,
  metadata: RecordNode | undefined,
): GroupBounds | undefined {
  const groupAtom = metadata && findDirect(metadata, RecordType.escherShapeGroup);
  if (!groupAtom || groupAtom.length < 16) return undefined;
  const bounds = {
    left: readInt32(view, groupAtom, 0),
    top: readInt32(view, groupAtom, 4),
    right: readInt32(view, groupAtom, 8),
    bottom: readInt32(view, groupAtom, 12),
  };
  if (groupAtom.length < 32) {
    return {
      ...bounds,
      childLeft: bounds.left,
      childTop: bounds.top,
      childRight: bounds.right,
      childBottom: bounds.bottom,
    };
  }
  return {
    ...bounds,
    childLeft: readInt32(view, groupAtom, 16),
    childTop: readInt32(view, groupAtom, 20),
    childRight: readInt32(view, groupAtom, 24),
    childBottom: readInt32(view, groupAtom, 28),
  };
}

function readTableFrame(view: DataView, metadata: RecordNode | undefined): TableFrame | undefined {
  const bounds = readGroupBounds(view, metadata);
  if (!bounds) return undefined;
  const shapeRecord = metadata && findDirect(metadata, RecordType.escherShape);
  return {
    ...(shapeRecord && shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };
}

function tableCell(shape: ShapeOptions): TableCell {
  const paragraphs = shape.textBody?.paragraphs;
  if (paragraphs !== undefined && paragraphs.length > 1) return { children: paragraphs };
  return { text: textBodyToText(shape.textBody) };
}

function parseShape(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): ShapeOptions | PictureOptions | undefined {
  const shapeRecord = findDirect(container, RecordType.escherShape);
  if (!shapeRecord) return undefined;
  const bounds = readAnchorBounds(view, container);
  if (!bounds) return undefined;
  const picture =
    shapeRecord.instance === 75 ? readShapePicture(view, container, context.pictures) : undefined;
  const text = readContainerText(view, container, context);
  if (!text && !picture) return undefined;

  const position = {
    ...(shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };
  if (picture && !text) return { ...position, type: picture.type, data: picture.data };
  if (picture) {
    return {
      ...position,
      textBox: true,
      properties: {
        fill: { type: "blip", data: picture.data, imageType: picture.type },
      },
      textBody: text,
    };
  }
  return {
    ...position,
    textBox: true,
    textBody: text,
  };
}

function readShapePicture(
  view: DataView,
  container: RecordNode,
  pictures: readonly (LegacyPicture | undefined)[],
): LegacyPicture | undefined {
  const options = findDirect(container, RecordType.escherShapeProperties);
  if (!options) return undefined;
  const reference = findShapePictureReference(view, options);
  if (reference === undefined || reference <= 0 || reference >= pictures.length) return undefined;
  return pictures[reference];
}

function findShapePictureReference(view: DataView, options: RecordNode): number | undefined {
  const headerEnd = options.offset + 8;
  const propertyCount = options.instance;
  const propertiesEnd = headerEnd + propertyCount * 6;
  if (propertiesEnd > options.end) return undefined;
  let cursor = options.offset + 8;
  let result: number | undefined;
  for (let index = 0; index < propertyCount; index += 1) {
    const packedId = view.getUint16(cursor, true);
    const value = view.getUint32(cursor + 2, true);
    const propertyId = packedId & 0x3fff;
    const isBlip = (packedId & 0x4000) !== 0;
    if (propertyId === 0x0104 && isBlip) result = value;
    cursor += 6;
  }
  return result;
}

function readContainerText(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): TextBodyOptions | undefined {
  const clientTextbox = findDirect(container, RecordType.escherClientTextbox);
  if (!clientTextbox) return undefined;

  const embeddedText = collectEmbeddedText(view, clientTextbox);
  if (embeddedText.length > 0) {
    const body = createTextBody(embeddedText);
    return applyClientHyperlinks(body, view, clientTextbox, context.hyperlinks);
  }

  const reference = findDescendant(clientTextbox, RecordType.outlineTextReference);
  if (reference && reference.length >= 4) {
    const index = readInt32(view, reference, 0);
    const entry = context.entries[index];
    if (entry) {
      context.usedEntryIndexes.add(index);
      return createTextBody([entry.text]);
    }
  }
  return undefined;
}

function readHyperlinkTargets(
  view: DataView,
  document: RecordNode,
): ReadonlyMap<number, { url?: string; tooltip?: string }> {
  const targets = new Map<number, { url?: string; tooltip?: string }>();
  for (const list of collectDescendants(document, RecordType.externalObjectList)) {
    for (const link of list.children.filter((child) => child.type === RecordType.hyperlink)) {
      const atom = findDirect(link, RecordType.hyperlinkAtom);
      const strings = link.children.filter((child) => child.type === RecordType.interactiveString);
      const id = atom && atom.length >= 4 ? readInt32(view, atom, 0) : undefined;
      const url = strings[1] ? decodeUtf16(recordBody(view, strings[1])) : undefined;
      const tooltip = strings[0] ? decodeUtf16(recordBody(view, strings[0])) : undefined;
      if (id === undefined || !url || url.includes(",")) continue;
      targets.set(id, { url, tooltip });
    }
  }
  return targets;
}

function applyClientHyperlinks(
  body: TextBodyOptions,
  view: DataView,
  clientTextbox: RecordNode,
  targets: ReadonlyMap<number, { url?: string; tooltip?: string }>,
): TextBodyOptions {
  const links: LegacyTextHyperlink[] = [];
  for (const [index, record] of clientTextbox.children.entries()) {
    if (record.type !== RecordType.interactiveInfo) continue;
    const atom = findDirect(record, RecordType.interactiveInfoAtom);
    const range = clientTextbox.children[index + 1];
    if (
      !atom ||
      atom.length < 8 ||
      !range ||
      range.type !== RecordType.interactiveText ||
      range.length < 8
    )
      continue;
    const target = targets.get(readInt32(view, atom, 4));
    if (!target?.url) continue;
    links.push({
      start: readInt32(view, range, 0),
      end: readInt32(view, range, 4),
      target,
    });
  }
  if (links.length === 0) return body;

  let position = 0;
  for (const paragraph of body.paragraphs ?? []) {
    if (typeof paragraph === "string") continue;
    for (const run of paragraph.children ?? []) {
      if (typeof run === "string" || (!("text" in run) && !("hyperlink" in run))) continue;
      const textRun = run as TextRunOptions;
      const start = position;
      const end = position + (textRun.text?.length ?? 0);
      const link = links.find((candidate) => start < candidate.end && end > candidate.start);
      if (link) {
        textRun.hyperlink = { url: link.target.url, tooltip: link.target.tooltip };
      }
      position = end;
    }
    position += 1;
  }
  return body;
}

function collectEmbeddedText(view: DataView, parent: RecordNode): string[] {
  const text: string[] = [];
  let placeholderType: number | undefined;
  let value: string | undefined;
  let hasDateTime = false;
  let hasSlideNumber = false;

  const flush = (): void => {
    if (value === undefined) return;
    if (value === "*" && hasSlideNumber) text.push("1");
    else if (value === "*" && hasDateTime) text.push("1/1/1");
    else text.push(value);
    value = undefined;
  };

  for (const child of flatten(parent)) {
    if (child.type === RecordType.textHeader) {
      flush();
      hasDateTime = false;
      hasSlideNumber = false;
      placeholderType = readInt32(view, child, 0);
    } else if (placeholderType !== undefined && child.type === RecordType.textChars) {
      value = decodeUtf16(recordBody(view, child));
    } else if (placeholderType !== undefined && child.type === RecordType.textBytes) {
      value = decodeAnsi(recordBody(view, child));
    } else if (child.type === RecordType.dateTimeAtom) {
      hasDateTime = true;
    } else if (child.type === RecordType.slideNumberAtom) {
      hasSlideNumber = true;
    }
  }
  flush();
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
  return value.split(/\r\n|\r/).map((paragraph) => paragraph.replace(/[\n\v]/g, "\n"));
}

function masterUnitsToEmu(value: number): number {
  return Math.round((value * EMU_PER_INCH) / MASTER_UNITS_PER_INCH);
}

import { CompoundFileReader } from "@office-open/core";
import type { TextBodyOptions, TextRunOptions } from "@office-open/core";
import type {
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
const PICTURE_RECORD_TYPES = new Set([
  0xf01a, 0xf01b, 0xf01c, 0xf01d, 0xf01e, 0xf01f, 0xf029, 0xf02a,
]);

interface TextEntry {
  readonly placeholderType: number;
  readonly text: string;
}

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

  const slideGroups = readSlideTextGroups(documentView, documentRecord, "slide");
  const masterGroups = readSlideTextGroups(documentView, documentRecord, "master");
  const masterNames = new Map<number, string>();
  const masters = masterGroups.map((group, index) => {
    const name = `Legacy Master ${index + 1}`;
    if (group.slideIdentifier !== undefined && !masterNames.has(group.slideIdentifier)) {
      masterNames.set(group.slideIdentifier, name);
    }
    return readMaster(documentView, persistReferences, group, pictures, name);
  });
  const defaultMaster = masters[0]?.name;
  const slides = slideGroups.map((group) =>
    readSlide(documentView, persistReferences, group, pictures, defaultMaster, masterNames),
  );
  const result: PresentationOptions = {};
  if (size) result.size = size;
  if (masters.length > 0) result.masters = masters;
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

function readSlideTextGroups(
  view: DataView,
  document: RecordNode,
  kind: "slide" | "master",
): SlideTextGroup[] {
  const lists = document.children
    .filter(
      (child) =>
        child.type === RecordType.slideListWithText &&
        child.instance === (kind === "slide" ? 0 : 1),
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
  pictures: readonly (LegacyPicture | undefined)[],
  defaultMaster: string | undefined,
  masterNames: ReadonlyMap<number, string>,
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
  const children = drawing
    ? readDrawingChildren(view, drawing, { entries: group.entries, pictures })
    : [];
  const slideAtom = findDirect(slide, RecordType.slideAtom);
  const masterId =
    slideAtom && slideAtom.length >= 16 ? readUint32(view, slideAtom, 12) : undefined;
  const master = masterId !== undefined && masterId > 0 ? masterNames.get(masterId) : undefined;
  const result: SlideOptions = {};
  if (children.length > 0) result.children = children;
  const slideMaster = master ?? defaultMaster;
  if (slideMaster) result.master = slideMaster;
  return result;
}

function readMaster(
  view: DataView,
  references: Map<number, number>,
  group: SlideTextGroup,
  pictures: readonly (LegacyPicture | undefined)[],
  name: string,
): MasterDefinition {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: missing master persist object",
    );
  }
  const master = readRecordTree(view, offset, view.byteLength)[0];
  if (!master || (master.type !== RecordType.slide && master.type !== RecordType.mainMaster)) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected master persist object",
    );
  }
  const drawing = findDirect(master, RecordType.ppDrawing);
  const children = drawing
    ? readDrawingChildren(view, drawing, { entries: group.entries, pictures })
    : [];
  return children.length > 0 ? { name, children } : { name };
}

function readPictureStore(data: Uint8Array): readonly (LegacyPicture | undefined)[] {
  const view = createView(data);
  const records = readRecordTree(view, 0, data.byteLength);
  const pictureRecords =
    records.length === 1 && records[0]!.type === RecordType.escherBStoreContainer
      ? records[0]!.children
      : records;
  if (pictureRecords.some((record) => !isPictureRecord(record))) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: malformed Pictures stream");
  }
  return [
    undefined,
    ...pictureRecords.map((record) =>
      record.type === RecordType.escherBse
        ? readPictureRecord(view, record)
        : readEmbeddedPicture(record, recordBody(view, record)),
    ),
  ];
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
  if (depth === 0) return children;

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
  if (lefts.length < 3 || tops.length < 2 || lefts.length * tops.length !== bounds.length) {
    return undefined;
  }

  const columnWidths: number[] = [];
  for (const [index, left] of lefts.entries()) {
    const cells = bounds.filter((item) => near(item.left, left));
    const right = cells[0]?.right;
    if (right === undefined || cells.some((item) => !near(item.right, right))) return undefined;
    const nextLeft = lefts[index + 1];
    if (nextLeft !== undefined && !near(right, nextLeft)) return undefined;
    const width = masterUnitsToEmu(right - left);
    if (width <= 0) return undefined;
    columnWidths.push(width);
  }

  const rows: NonNullable<TableOptions["rows"]> = [];
  for (const [index, top] of tops.entries()) {
    const cells = bounds
      .filter((item) => near(item.top, top))
      .sort((first, second) => first.left - second.left);
    if (cells.length !== lefts.length) return undefined;
    const bottom = cells[0]!.bottom;
    if (cells.some((item) => !near(item.bottom, bottom))) return undefined;
    const nextTop = tops[index + 1];
    if (nextTop !== undefined && !near(bottom, nextTop)) return undefined;
    const height = masterUnitsToEmu(bottom - top);
    if (height <= 0) return undefined;
    rows.push({
      height,
      cells: cells.map((cell) => ({ text: textBodyToText(cell.shape.textBody) })),
    });
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
  if (!groupAtom || groupAtom.length < 32) return undefined;
  return {
    left: readInt32(view, groupAtom, 0),
    top: readInt32(view, groupAtom, 4),
    right: readInt32(view, groupAtom, 8),
    bottom: readInt32(view, groupAtom, 12),
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

function parseShape(
  view: DataView,
  container: RecordNode,
  context: DrawingContext,
): ShapeOptions | PictureOptions | undefined {
  const shapeRecord = findDirect(container, RecordType.escherShape);
  if (!shapeRecord) return undefined;
  const bounds = readAnchorBounds(view, container);
  if (!bounds) return undefined;
  const text = readContainerText(view, container, context.entries);
  const picture =
    shapeRecord.instance === 75 ? readShapePicture(view, container, context.pictures) : undefined;
  if (!text && !picture) return undefined;

  const position = {
    ...(shapeRecord.length >= 4 ? { id: readInt32(view, shapeRecord, 0) } : {}),
    x: masterUnitsToEmu(Math.min(bounds.left, bounds.right)),
    y: masterUnitsToEmu(Math.min(bounds.top, bounds.bottom)),
    width: masterUnitsToEmu(Math.abs(bounds.right - bounds.left)),
    height: masterUnitsToEmu(Math.abs(bounds.bottom - bounds.top)),
  };
  if (picture) return { ...position, type: picture.type, data: picture.data };
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

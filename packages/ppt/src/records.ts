const RECORD_HEADER_SIZE = 8;

export const RecordType = {
  document: 1000,
  documentAtom: 1001,
  slide: 1006,
  slideAtom: 1007,
  mainMaster: 1016,
  slidePersistAtom: 1011,
  ppDrawing: 1036,
  textHeader: 3999,
  textChars: 4000,
  textBytes: 4008,
  slideListWithText: 4080,
  userEditAtom: 4085,
  currentUserAtom: 4086,
  persistPointerFullBlock: 6001,
  persistPointerIncrementalBlock: 6002,
  documentEncryptionAtom: 12052,
  outlineTextReference: 3998,
  escherContainer: 0xf000,
  escherBStoreContainer: 0xf001,
  escherDrawingContainer: 0xf002,
  escherShapeGroupContainer: 0xf003,
  escherShapeContainer: 0xf004,
  escherBse: 0xf007,
  escherShapeGroup: 0xf009,
  escherShape: 0xf00a,
  escherShapeProperties: 0xf00b,
  escherClientTextbox: 0xf00d,
  escherChildAnchor: 0xf00f,
  escherClientAnchor: 0xf010,
} as const;

export interface RecordHeader {
  readonly offset: number;
  readonly version: number;
  readonly instance: number;
  readonly type: number;
  readonly length: number;
  readonly end: number;
}

export interface RecordNode extends RecordHeader {
  readonly children: readonly RecordNode[];
}

export class LegacyPowerPointError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "LegacyPowerPointError";
  }
}

export function isContainer(record: RecordHeader): boolean {
  return record.version === 15;
}

export function readRecordHeader(
  view: DataView,
  offset: number,
  streamEnd = view.byteLength,
): RecordHeader {
  if (offset < 0 || offset + RECORD_HEADER_SIZE > streamEnd) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint record: truncated record header");
  }
  const packed = view.getUint16(offset, true);
  const length = view.getUint32(offset + 4, true);
  const end = offset + RECORD_HEADER_SIZE + length;
  if (end > streamEnd) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint record: record exceeds its parent");
  }
  return {
    offset,
    version: packed & 0xf,
    instance: packed >>> 4,
    type: view.getUint16(offset + 2, true),
    length,
    end,
  };
}

export function readRecordTree(
  view: DataView,
  offset: number,
  end: number,
  depth = 0,
): RecordNode[] {
  const maxDepth = 64;
  const maxRecords = 100_000;
  const nodes: RecordNode[] = [];
  let cursor = offset;

  while (cursor + RECORD_HEADER_SIZE <= end) {
    if (isZeroPadding(view, cursor, end)) break;
    if (nodes.length >= maxRecords) {
      throw new LegacyPowerPointError("Corrupt legacy PowerPoint record: too many sibling records");
    }
    const header = readRecordHeader(view, cursor, end);
    const children =
      isContainer(header) && header.length > 0
        ? depth < maxDepth
          ? readRecordTree(view, header.offset + RECORD_HEADER_SIZE, header.end, depth + 1)
          : readAtomChildren(view, header)
        : [];
    nodes.push({ ...header, children });
    cursor = header.end;
  }

  if (cursor !== end && !isZeroPadding(view, cursor, end)) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint record: trailing bytes in container",
    );
  }
  return nodes;
}

function isZeroPadding(view: DataView, start: number, end: number): boolean {
  for (let offset = start; offset < end; offset += 1) {
    if (view.getUint8(offset) !== 0) return false;
  }
  return true;
}

function readAtomChildren(view: DataView, parent: RecordHeader): RecordNode[] {
  const children: RecordNode[] = [];
  let cursor = parent.offset + RECORD_HEADER_SIZE;
  while (cursor + RECORD_HEADER_SIZE <= parent.end) {
    const header = readRecordHeader(view, cursor, parent.end);
    children.push({ ...header, children: [] });
    cursor = header.end;
  }
  return children;
}

export function findDirect(record: RecordNode, type: number): RecordNode | undefined {
  return record.children.find((child) => child.type === type);
}

export function findAllDirect(records: readonly RecordNode[], type: number): RecordNode[] {
  return records.filter((record) => record.type === type);
}

export function findDescendant(record: RecordNode, type: number): RecordNode | undefined {
  for (const child of record.children) {
    if (child.type === type) return child;
    const descendant = findDescendant(child, type);
    if (descendant) return descendant;
  }
  return undefined;
}

export function collectDescendants(
  record: RecordNode,
  type: number,
  output: RecordNode[] = [],
): RecordNode[] {
  for (const child of record.children) {
    if (child.type === type) output.push(child);
    collectDescendants(child, type, output);
  }
  return output;
}

export function readInt32(view: DataView, record: RecordHeader, offset: number): number {
  if (offset < 0 || offset + 4 > record.length) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint atom: missing 32-bit field");
  }
  return view.getInt32(record.offset + RECORD_HEADER_SIZE + offset, true);
}

export function readUint32(view: DataView, record: RecordHeader, offset: number): number {
  if (offset < 0 || offset + 4 > record.length) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint atom: missing 32-bit field");
  }
  return view.getUint32(record.offset + RECORD_HEADER_SIZE + offset, true);
}

export function readInt16(view: DataView, record: RecordHeader, offset: number): number {
  if (offset < 0 || offset + 2 > record.length) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint atom: missing 16-bit field");
  }
  return view.getInt16(record.offset + RECORD_HEADER_SIZE + offset, true);
}

export function recordBody(view: DataView, record: RecordHeader): Uint8Array {
  return new Uint8Array(
    view.buffer,
    view.byteOffset + record.offset + RECORD_HEADER_SIZE,
    record.length,
  );
}

export function decodeAnsi(bytes: Uint8Array): string {
  let result = "";
  for (const byte of bytes) result += String.fromCharCode(byte);
  return result;
}

export function decodeUtf16(bytes: Uint8Array): string {
  const parity = bytes.byteLength % 2;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength - parity);
  let result = "";
  for (let offset = 0; offset < view.byteLength; offset += 2) {
    result += String.fromCharCode(view.getUint16(offset, true));
  }
  return result;
}

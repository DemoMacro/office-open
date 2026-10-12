import type { ShapeOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";
import { RecordType, findDirect, type RecordNode } from "../records";

type PlaceholderType = NonNullable<ShapeOptions["placeholder"]>;
type PlaceholderSize = NonNullable<ShapeOptions["placeholderSize"]>;

interface PlaceholderAtom {
  readonly index: number;
  readonly placeholder: PlaceholderType | undefined;
  readonly size: PlaceholderSize;
}

const PLACEHOLDER_TYPES: readonly (PlaceholderType | undefined)[] = [
  undefined,
  "title",
  "body",
  "centerTitle",
  "subTitle",
  "slideImage",
  "body",
  "date",
  "slideNumber",
  "footer",
  "header",
  "slideImage",
  "body",
  "title",
  "body",
  "centerTitle",
  "subTitle",
  "title",
  "body",
  "object",
  "chart",
  "table",
  "clipArt",
  undefined,
  "media",
  "object",
  "picture",
];

export interface ShapePlaceholder {
  readonly placeholder?: PlaceholderType;
  readonly placeholderIndex?: number;
  readonly placeholderSize?: PlaceholderSize;
}

export function readShapePlaceholder(
  view: DataView,
  container: RecordNode,
): ShapePlaceholder | null | undefined {
  const clientData = findDirect(container, RecordType.escherClientData);
  const atom = clientData && findDirect(clientData, RecordType.placeholder);
  const headerFooterAtom =
    clientData && findDirect(clientData, RecordType.roundTripHeaderFooterPlaceholder);
  const newPlaceholderAtom =
    clientData && findDirect(clientData, RecordType.roundTripNewPlaceholder);
  if (!atom && !headerFooterAtom && !newPlaceholderAtom) return undefined;

  const parsedAtom = atom ? parsePlaceholderAtom(view, atom) : undefined;
  if (parsedAtom === null) return null;
  if (headerFooterAtom && headerFooterAtom.length !== 1) {
    throw invalidPlaceholderAtom(
      headerFooterAtom,
      "invalid-placeholder-atom-length",
      "invalid length",
    );
  }
  if (newPlaceholderAtom && newPlaceholderAtom.length !== 1) {
    throw invalidPlaceholderAtom(
      newPlaceholderAtom,
      "invalid-placeholder-atom-length",
      "invalid length",
    );
  }

  const placeholder =
    newPlaceholderAtom && newPlaceholderAtom.length >= 1
      ? placeholderType(view.getUint8(newPlaceholderAtom.offset + 8))
      : (parsedAtom?.placeholder ??
        (headerFooterAtom && headerFooterAtom.length >= 1
          ? placeholderType(view.getUint8(headerFooterAtom.offset + 8))
          : undefined));

  return {
    ...(placeholder !== undefined ? { placeholder } : {}),
    ...(parsedAtom ? { placeholderIndex: parsedAtom.index } : {}),
    ...(parsedAtom && parsedAtom.size !== "full" ? { placeholderSize: parsedAtom.size } : {}),
  };
}

function parsePlaceholderAtom(view: DataView, record: RecordNode): PlaceholderAtom | null {
  if (record.length !== 8) {
    throw invalidPlaceholderAtom(record, "invalid-placeholder-atom-length", "invalid length");
  }

  const index = view.getInt32(record.offset + 8, true);
  if (index === -1) return null;
  if (index < 0) {
    throw invalidPlaceholderAtom(record, "invalid-placeholder-index", "invalid index");
  }

  return {
    index,
    placeholder: placeholderType(view.getUint8(record.offset + 12)),
    size: placeholderSize(view.getUint8(record.offset + 13)),
  };
}

function placeholderType(value: number): PlaceholderType | undefined {
  if (value === 0 || value >= PLACEHOLDER_TYPES.length) {
    throw new PptParseError("Corrupt legacy PowerPoint placeholder atom: unknown type", {
      part: "record",
      recordType: RecordType.placeholder,
      recordName: "PlaceholderAtom",
      reason: "unknown-placeholder-type",
    });
  }
  return PLACEHOLDER_TYPES[value]!;
}

function placeholderSize(value: number): PlaceholderSize {
  if (value > 2) {
    throw new PptParseError("Corrupt legacy PowerPoint placeholder atom: unknown size", {
      part: "record",
      recordType: RecordType.placeholder,
      recordName: "PlaceholderAtom",
      reason: "unknown-placeholder-size",
    });
  }
  return value === 1 ? "half" : value === 2 ? "quarter" : "full";
}

function invalidPlaceholderAtom(
  record: RecordNode,
  reason: string,
  description: string,
): PptParseError {
  return new PptParseError(`Corrupt legacy PowerPoint placeholder atom: ${description}`, {
    part: "record",
    recordType: record.type,
    recordName: "PlaceholderAtom",
    offset: record.offset + 8,
    length: 8,
    byteRange: [record.offset + 8, record.end],
    reason,
  });
}

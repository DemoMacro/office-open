import type { CorePropertiesOptions } from "../opc/core";

const VT_I2 = 2;
const VT_I4 = 3;
const VT_BOOL = 11;
const VT_LPSTR = 30;
const VT_LPWSTR = 31;
const VT_FILETIME = 64;

const SUMMARY_IDS = {
  codepage: 1,
  title: 2,
  subject: 3,
  author: 4,
  keywords: 5,
  comments: 6,
  lastAuthor: 8,
  revision: 9,
  lastPrinted: 13,
  created: 15,
  modified: 18,
} as const;

const DOCUMENT_SUMMARY_IDS = {
  codepage: 1,
  category: 2,
} as const;

function view(data: Uint8Array): DataView {
  return new DataView(data.buffer, data.byteOffset, data.byteLength);
}

function decodeString(
  data: Uint8Array,
  view: DataView,
  offset: number,
  wide: boolean,
  codepage: number,
): string {
  if (offset + 4 > data.byteLength) return "";
  const characters = view.getUint32(offset, true);
  if (wide) {
    const bytes = Math.min(characters * 2, data.byteLength - offset - 4);
    return new TextDecoder("utf-16le")
      .decode(data.subarray(offset + 4, offset + 4 + bytes))
      .replaceAll("\0", "");
  }
  const bytes = Math.min(characters, data.byteLength - offset - 4);
  const encoded = data.subarray(offset + 4, offset + 4 + bytes);
  if (codepage === 65001) {
    return new TextDecoder("utf-8").decode(encoded).replaceAll("\0", "");
  }
  return Array.from(
    encoded.filter((byte) => byte !== 0),
    (byte) => String.fromCharCode(byte),
  ).join("");
}

function fileTimeToW3cdtf(view: DataView, offset: number): string | undefined {
  if (offset + 8 > view.byteLength) return undefined;
  const low = view.getUint32(offset, true);
  const high = view.getUint32(offset + 4, true);
  const value = high * 2 ** 32 + low;
  if (value === 0) return undefined;
  const milliseconds = Math.floor((value - 116444736000000000) / 10_000);
  const date = new Date(milliseconds);
  if (Number.isNaN(date.getTime())) return undefined;
  return date.toISOString().replace(/\.\d+Z$/, "Z");
}

function parseProperties(
  data: Uint8Array,
  section: number,
  ids: Record<string, number>,
): Record<string, string | number | boolean | undefined> {
  const propertyView = view(data);
  if (section + 8 > data.byteLength) return {};
  const count = propertyView.getUint32(section + 4, true);
  const values: Record<string, string | number | boolean | undefined> = {};
  let codepage = 1252;

  for (let index = 0; index < count; index += 1) {
    const entry = section + 8 + index * 8;
    if (entry + 8 > data.byteLength) break;
    const id = propertyView.getUint32(entry, true);
    const valueOffset = section + propertyView.getUint32(entry + 4, true);
    if (valueOffset + 4 > data.byteLength) continue;
    const type = propertyView.getUint16(valueOffset, true);
    const body = valueOffset + 4;
    if (id === ids.codepage && type === VT_I2 && body + 2 <= data.byteLength) {
      codepage = propertyView.getUint16(body, true);
      continue;
    }
    if (type === VT_LPSTR)
      values[String(id)] = decodeString(data, propertyView, body, false, codepage);
    else if (type === VT_LPWSTR)
      values[String(id)] = decodeString(data, propertyView, body, true, 1200);
    else if (type === VT_I2 && body + 2 <= data.byteLength)
      values[String(id)] = propertyView.getInt16(body, true);
    else if (type === VT_I4 && body + 4 <= data.byteLength)
      values[String(id)] = propertyView.getInt32(body, true);
    else if (type === VT_BOOL && body + 2 <= data.byteLength)
      values[String(id)] = propertyView.getUint16(body, true) !== 0;
    else if (type === VT_FILETIME) {
      const time = fileTimeToW3cdtf(propertyView, body);
      if (time !== undefined) values[String(id)] = time;
    }
  }
  return values;
}

function property(
  values: Record<string, string | number | boolean | undefined>,
  id: number,
): string | number | boolean | undefined {
  return values[String(id)];
}

function text(value: string | number | boolean | undefined): string | undefined {
  return value === undefined || typeof value === "boolean" ? undefined : String(value);
}

/** Read the standard and document summary property sets from a CFB stream. */
export function parseSummaryInformation(data: Uint8Array): CorePropertiesOptions {
  if (data.byteLength < 48) return {};
  const summaryView = view(data);
  const summaryOffset = summaryView.getUint32(44, true);
  const summary = parseProperties(data, summaryOffset, SUMMARY_IDS);
  const propertySetCount = summaryView.getUint32(24, true);
  const documentOffset = propertySetCount >= 2 ? summaryView.getUint32(64, true) : undefined;
  const documentSummary =
    documentOffset !== undefined && Number.isSafeInteger(documentOffset)
      ? parseProperties(data, documentOffset, DOCUMENT_SUMMARY_IDS)
      : {};
  const created = property(summary, SUMMARY_IDS.created);
  const modified = property(summary, SUMMARY_IDS.modified);
  const revision = property(summary, SUMMARY_IDS.revision);

  return {
    ...(text(property(summary, SUMMARY_IDS.title)) !== undefined
      ? { title: text(property(summary, SUMMARY_IDS.title)) }
      : {}),
    ...(text(property(summary, SUMMARY_IDS.subject)) !== undefined
      ? { subject: text(property(summary, SUMMARY_IDS.subject)) }
      : {}),
    ...(text(property(summary, SUMMARY_IDS.author)) !== undefined
      ? { creator: text(property(summary, SUMMARY_IDS.author)) }
      : {}),
    ...(text(property(summary, SUMMARY_IDS.keywords)) !== undefined
      ? { keywords: text(property(summary, SUMMARY_IDS.keywords)) }
      : {}),
    ...(text(property(summary, SUMMARY_IDS.comments)) !== undefined
      ? { description: text(property(summary, SUMMARY_IDS.comments)) }
      : {}),
    ...(text(property(summary, SUMMARY_IDS.lastAuthor)) !== undefined
      ? { lastModifiedBy: text(property(summary, SUMMARY_IDS.lastAuthor)) }
      : {}),
    ...(created === undefined ? {} : { created: String(created) }),
    ...(modified === undefined ? {} : { modified: String(modified) }),
    ...(typeof revision === "number" ? { revision } : {}),
    ...(text(property(summary, SUMMARY_IDS.lastPrinted)) !== undefined
      ? { lastPrinted: text(property(summary, SUMMARY_IDS.lastPrinted)) }
      : {}),
    ...(text(property(documentSummary, DOCUMENT_SUMMARY_IDS.category)) !== undefined
      ? { category: text(property(documentSummary, DOCUMENT_SUMMARY_IDS.category)) }
      : {}),
  };
}

/** Read the category-only document summary property set from a CFB stream. */
export function parseDocumentSummaryInformation(data: Uint8Array): CorePropertiesOptions {
  if (data.byteLength < 48) return {};
  const section = view(data).getUint32(44, true);
  const values = parseProperties(data, section, DOCUMENT_SUMMARY_IDS);
  const category = property(values, DOCUMENT_SUMMARY_IDS.category);
  return category === undefined ? {} : { category: text(category) };
}

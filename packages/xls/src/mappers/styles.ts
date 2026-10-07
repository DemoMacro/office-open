import type { FontOptions } from "@office-open/xlsx";

import { decodeBiffString, readBiff5String, readBiff8String } from "./biff";
export function parseFont(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
): FontOptions | undefined {
  if (body.byteLength < 15) return undefined;
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  // MS-XLS 2.5.146: Font.name is a shortXLUnicodeString — a 1-byte cch (not
  // the 2-byte XLUnicodeString count), preceded by 14 fixed bytes.
  let fontName = "";
  try {
    const characterCount = body[14]!;
    fontName = isBiff8
      ? decodeBiffString(body, view, 16, characterCount, body[15] ?? 0, encoding).value
      : new TextDecoder(encoding).decode(body.subarray(15, 15 + characterCount));
  } catch {
    fontName = "";
  }
  const weight = view.getUint16(6, true);
  return {
    size: view.getUint16(0, true) / 20,
    bold: weight >= 700,
    italic: (view.getUint16(2, true) & 0x0002) !== 0,
    color: { indexed: view.getUint16(4, true) },
    underline: body[10] === 1,
    font: fontName,
  };
}

export function parseFormat(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
): { id: number; code: string } | undefined {
  if (body.byteLength < 4) return undefined;
  const id = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true);
  let code: string;
  try {
    code = (isBiff8 ? readBiff8String(body, 2, encoding) : readBiff5String(body, 2, encoding))
      .value;
  } catch {
    return undefined;
  }
  return {
    id,
    code,
  };
}

export function parsePalette(body: Uint8Array): string[] {
  if (body.byteLength < 2) return [];
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const count = view.getUint16(0, true);
  if (body.byteLength < 2 + count * 4) return [];
  return Array.from({ length: count }, (_, index) => {
    const value = view.getUint32(2 + index * 4, true);
    const blue = (value & 0xff).toString(16).padStart(2, "0");
    const green = ((value >> 8) & 0xff).toString(16).padStart(2, "0");
    const red = ((value >> 16) & 0xff).toString(16).padStart(2, "0");
    return `FF${red}${green}${blue}`;
  });
}

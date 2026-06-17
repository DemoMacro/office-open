/**
 * Output type definitions for OOXML document export.
 *
 * @module
 */

import { strFromU8 } from "fflate";

export interface OutputByType {
  base64: string;
  string: string;
  text: string;
  binarystring: string;
  array: readonly number[];
  uint8array: Uint8Array;
  arraybuffer: ArrayBuffer;
  blob: Blob;
  nodebuffer: Buffer;
}

export type OutputType = keyof OutputByType;

/* V8 ignore start */
export const OoxmlMimeType = {
  DOCX: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  PPTX: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  XLSX: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
} as const;

export const convertOutput = <T extends OutputType>(
  data: Uint8Array,
  type: T,
  mimeType: string = OoxmlMimeType.DOCX,
): OutputByType[T] => {
  switch (type) {
    case "nodebuffer":
      return Buffer.from(data) as OutputByType[T];
    case "blob":
      return new Blob(
        [(data.buffer as ArrayBuffer).slice(data.byteOffset, data.byteOffset + data.byteLength)],
        { type: mimeType },
      ) as OutputByType[T];
    case "arraybuffer":
      return data.buffer.slice(
        data.byteOffset,
        data.byteOffset + data.byteLength,
      ) as OutputByType[T];
    case "uint8array":
      return data as OutputByType[T];
    case "base64":
      // Use Buffer in Node.js (more robust), btoa as browser fallback
      return (
        typeof Buffer !== "undefined"
          ? Buffer.from(data).toString("base64")
          : btoa(strFromU8(data, true))
      ) as OutputByType[T];
    case "string":
    case "text":
    case "binarystring":
      return strFromU8(data, true) as OutputByType[T];
    case "array":
      return [...data] as unknown as OutputByType[T];
    /* V8 ignore next */
    default:
      return data as unknown as OutputByType[T];
  }
};
/* V8 ignore stop */

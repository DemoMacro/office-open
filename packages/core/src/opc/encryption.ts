/**
 * Encrypted OOXML container detection and passthrough helpers.
 *
 * Password-protected OOXML files (ECMA-376 agile/standard encryption) wrap
 * the OPC package inside an OLE2/CFB compound file. The plaintext is not
 * recoverable without the password, so parse() carries the original bytes on
 * the options and generate() re-emits them verbatim instead of compiling.
 *
 * @module
 */

import type { DataType } from "../util/data-type";
import { toUint8Array } from "../util/data-type";
import type { OutputByType, OutputType } from "./output";
import { convertOutput } from "./output";

/** OLE2/CFB compound-file signature shared by every encrypted OOXML file. */
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

/**
 * Verbatim payload of a source file whose OOXML package is encrypted.
 *
 * Round-trip only: the plaintext needs the password, so these bytes are the
 * only faithful representation. Every other options field stays empty —
 * there is no content to model until the file is decrypted.
 */
export interface EncryptedContainerOptions {
  /** Complete source file bytes (OLE2/CFB container). */
  data: DataType;
}

/** Whether `data` is an OLE2/CFB container — i.e. an encrypted OOXML file. */
export function isEncryptedContainer(data: Uint8Array): boolean {
  return data.length >= CFB_SIGNATURE.length && CFB_SIGNATURE.every((b, i) => data[i] === b);
}

/**
 * Converted output for an encrypted-container file, or `undefined` when the
 * options carry no encrypted payload (the normal compile path applies).
 */
export function encryptedContainerOutput<T extends OutputType>(
  file: { encrypted?: EncryptedContainerOptions },
  type: T,
  mimeType: string,
): OutputByType[T] | undefined {
  const { encrypted } = file;
  if (!encrypted) return undefined;
  return convertOutput(toUint8Array(encrypted.data), type, mimeType);
}

/** Streaming variant of {@link encryptedContainerOutput}. */
export function encryptedContainerStream(file: {
  encrypted?: EncryptedContainerOptions;
}): ReadableStream<Uint8Array> | undefined {
  const { encrypted } = file;
  if (!encrypted) return undefined;
  const bytes = toUint8Array(encrypted.data);
  return new ReadableStream<Uint8Array>({
    start(c) {
      c.enqueue(bytes);
      c.close();
    },
  });
}

/**
 * Reject encrypted passthrough mixed with real content: generate() re-emits
 * the encrypted bytes, so every other field would be silently dropped.
 *
 * `hasContent` is the package's primary content collection being non-empty
 * (e.g. `sections.length > 0`) — callers know their own shape.
 */
export function assertEncryptedExclusive(
  file: { encrypted?: EncryptedContainerOptions },
  hasContent: boolean,
): void {
  if (file.encrypted && hasContent) {
    throw new Error(
      "Encrypted passthrough carries no editable content — clear the content fields or drop `encrypted`.",
    );
  }
}

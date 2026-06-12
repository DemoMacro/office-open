/**
 * Unique ID generation utilities.
 *
 * @module
 */

import { sha1 } from "@noble/hashes/legacy.js";
import { bytesToHex } from "@noble/hashes/utils.js";

/**
 * A function that generates unique sequential numeric IDs.
 */
export type UniqueNumericIdCreator = () => number;

/**
 * Creates a unique numeric ID generator with sequential numbering.
 */
export const uniqueNumericIdCreator = (initial = 0): UniqueNumericIdCreator => {
  let currentCount = initial;
  return () => ++currentCount;
};

const URL_ALPHABET = "useandom-26T198340PX75pxJACKVERYMINDBUSHWOLF_GQZbfghjklqvwyzrict";

/**
 * Generates a unique lowercase alphanumeric ID using crypto.getRandomValues.
 */
export const uniqueId = (): string => {
  const bytes = new Uint8Array(21);
  crypto.getRandomValues(bytes);
  let id = "";
  for (let i = 0; i < 21; i++) id += URL_ALPHABET[bytes[i] & 63];
  return id.toLowerCase();
};

/**
 * Generates a SHA-1 hash of the provided data.
 */
export const hashedId = (data: Uint8Array | ArrayBuffer | string): string => {
  const bytes =
    data instanceof ArrayBuffer
      ? new Uint8Array(data)
      : typeof data === "string"
        ? new TextEncoder().encode(data)
        : data;
  return bytesToHex(sha1(bytes));
};

/**
 * Generates a UUID v4-style unique identifier using crypto.randomUUID.
 */
export const uniqueUuid = (): string => crypto.randomUUID();

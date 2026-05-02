/**
 * Unique ID generation utilities.
 *
 * @module
 */
import hash from "hash.js";
import { customAlphabet, nanoid } from "nanoid/non-secure";

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

/**
 * Generates a unique lowercase alphanumeric ID using nanoid.
 */
export const uniqueId = (): string => nanoid().toLowerCase();

/**
 * Generates a SHA-1 hash of the provided data.
 */
export const hashedId = (data: Buffer | string | Uint8Array | ArrayBuffer): string =>
    hash
        .sha1()
        .update(data instanceof ArrayBuffer ? new Uint8Array(data) : data)
        .digest("hex");

/**
 * Generates a random hexadecimal string of specified length.
 */
const generateUuidPart = (count: number): string => customAlphabet("1234567890abcdef", count)();

/**
 * Generates a UUID v4-style unique identifier.
 */
export const uniqueUuid = (): string =>
    `${generateUuidPart(8)}-${generateUuidPart(4)}-${generateUuidPart(4)}-${generateUuidPart(4)}-${generateUuidPart(12)}`;

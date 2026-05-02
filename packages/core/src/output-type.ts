/**
 * Output type definitions for OOXML document export.
 *
 * @module
 */

export interface OutputByType {
    readonly base64: string;
    readonly string: string;
    readonly text: string;
    readonly binarystring: string;
    readonly array: readonly number[];
    readonly uint8array: Uint8Array;
    readonly arraybuffer: ArrayBuffer;
    readonly blob: Blob;
    readonly nodebuffer: Buffer;
}

export type OutputType = keyof OutputByType;

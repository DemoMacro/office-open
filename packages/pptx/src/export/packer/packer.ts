import { Readable } from "stream";

import type { File } from "@file/file";
import { type ZipOptions, Zip, ZipDeflate, ZipPassThrough, zipSync } from "fflate";

import { Compiler } from "./next-compiler";
import type { IXmlifyedFile } from "./next-compiler";

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

export const convertOutput = <T extends OutputType>(data: Uint8Array, type: T): OutputByType[T] => {
    switch (type) {
        case "nodebuffer":
            return Buffer.from(data) as OutputByType[T];
        case "blob":
            return new Blob(
                [
                    (data.buffer as ArrayBuffer).slice(
                        data.byteOffset,
                        data.byteOffset + data.byteLength,
                    ),
                ],
                {
                    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                },
            ) as OutputByType[T];
        case "arraybuffer":
            return data.buffer.slice(
                data.byteOffset,
                data.byteOffset + data.byteLength,
            ) as OutputByType[T];
        case "uint8array":
            return data as OutputByType[T];
        case "base64":
            return btoa(
                Array.from(data, (b) => String.fromCharCode(b)).join(""),
            ) as OutputByType[T];
        case "string":
        case "text":
        case "binarystring":
            return new TextDecoder().decode(data) as OutputByType[T];
        case "array":
            return [...data] as unknown as OutputByType[T];
        default:
            return data as unknown as OutputByType[T];
    }
};

export const PrettifyType = {
    NONE: "",
    WITH_2_BLANKS: "  ",
    WITH_4_BLANKS: "    ",
    WITH_TAB: "\t",
} as const;

const convertPrettifyType = (
    prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
): (typeof PrettifyType)[keyof typeof PrettifyType] | undefined =>
    prettify === true ? PrettifyType.WITH_2_BLANKS : prettify === false ? undefined : prettify;

export class Packer {
    public static async pack<T extends OutputType>(
        file: File,
        type: T,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<OutputByType[T]> {
        const files = this.compiler.compile(file, convertPrettifyType(prettify), overrides);
        const zipped = zipSync(files, { level: 6 });
        return convertOutput(zipped, type);
    }

    public static toBuffer(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<Buffer> {
        return Packer.pack(file, "nodebuffer", prettify, overrides);
    }

    public static toBlob(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<Blob> {
        return Packer.pack(file, "blob", prettify, overrides);
    }

    public static toBase64String(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<string> {
        return Packer.pack(file, "base64", prettify, overrides);
    }

    public static toStream(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Readable {
        const stream = new Readable({ read() {} });

        try {
            const files = this.compiler.compile(file, convertPrettifyType(prettify), overrides);

            const zip = new Zip((err, chunk, final) => {
                if (err) {
                    stream.destroy(err);
                    return;
                }
                if (!stream.destroyed) {
                    stream.push(chunk);
                }
                if (final) {
                    stream.push(null);
                }
            });

            for (const [name, data] of Object.entries(files)) {
                const raw = Array.isArray(data) ? (data[0] as Uint8Array) : (data as Uint8Array);
                const level = Array.isArray(data) ? ((data[1] as ZipOptions).level ?? 6) : 6;
                const entry =
                    level === 0 ? new ZipPassThrough(name) : new ZipDeflate(name, { level });
                zip.add(entry);
                entry.push(raw, true);
            }

            zip.end();
        } catch (err) {
            stream.destroy(err instanceof Error ? err : new Error(String(err)));
        }

        return stream;
    }

    private static readonly compiler = new Compiler();
}

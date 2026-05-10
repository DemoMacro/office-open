import { Readable } from "stream";

import type { File } from "@file/file";
import { convertOutput, convertPrettifyType, OoxmlMimeType, PrettifyType } from "@office-open/core";
export { PrettifyType } from "@office-open/core";
import type { IXmlifyedFile, OutputByType, OutputType } from "@office-open/core";
import { type ZipOptions, Zip, ZipDeflate, ZipPassThrough, zipSync } from "fflate";

import { Compiler } from "./next-compiler";

export class Packer {
    public static async pack<T extends OutputType>(
        file: File,
        type: T,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<OutputByType[T]> {
        const files = this.compiler.compile(file, convertPrettifyType(prettify), overrides);
        const zipped = zipSync(files, { level: 6 });
        return convertOutput(zipped, type, OoxmlMimeType.PPTX);
    }

    public static async toString(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<string> {
        return Packer.pack(file, "string", prettify, overrides);
    }

    public static toArrayBuffer(
        file: File,
        prettify?: boolean | (typeof PrettifyType)[keyof typeof PrettifyType],
        overrides: readonly IXmlifyedFile[] = [],
    ): Promise<ArrayBuffer> {
        return Packer.pack(file, "arraybuffer", prettify, overrides);
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

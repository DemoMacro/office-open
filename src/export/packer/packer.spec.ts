import { File } from "@file/file";
import { HeadingLevel, Paragraph } from "@file/paragraph";
import { afterEach, assert, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { Packer, PrettifyType } from "./packer";

describe("Packer", () => {
    let file: File;

    beforeEach(() => {
        file = new File({
            creator: "Dolan Miu",
            lastModifiedBy: "Dolan Miu",
            revision: 1,
            sections: [
                {
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.TITLE,
                            text: "title",
                        }),
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            text: "Hello world",
                        }),
                        new Paragraph({
                            heading: HeadingLevel.HEADING_2,
                            text: "heading 2",
                        }),
                        new Paragraph("test text"),
                    ],
                },
            ],
        });
    });

    describe("prettify", () => {
        afterEach(() => {
            vi.restoreAllMocks();
        });

        it("should use a default prettify value", async () => {
            const spy = vi.spyOn((Packer as any).compiler, "compile");

            await Packer.toString(file, true);

            expect(spy).toBeCalledWith(
                expect.anything(),
                PrettifyType.WITH_2_BLANKS,
                expect.anything(),
            );
        });

        it("should use a prettify value", async () => {
            const spy = vi.spyOn((Packer as any).compiler, "compile");

            await Packer.toString(file, PrettifyType.WITH_4_BLANKS);

            expect(spy).toBeCalledWith(
                expect.anything(),
                PrettifyType.WITH_4_BLANKS,
                expect.anything(),
            );
        });

        it("should use an undefined prettify value", async () => {
            const spy = vi.spyOn((Packer as any).compiler, "compile");

            await Packer.toString(file, false);

            expect(spy).toBeCalledWith(expect.anything(), undefined, expect.anything());
        });
    });

    describe("overrides", () => {
        afterEach(() => {
            vi.restoreAllMocks();
        });

        it("should use an overrides value", async () => {
            const spy = vi.spyOn((Packer as any).compiler, "compile");
            const overrides = [{ data: "comments", path: "word/comments.xml" }];

            await Packer.toString(file, true, overrides);

            expect(spy).toBeCalledWith(expect.anything(), expect.anything(), overrides);
        });

        it("should use a default overrides value", async () => {
            const spy = vi.spyOn((Packer as any).compiler, "compile");

            await Packer.toString(file);

            expect(spy).toBeCalledWith(expect.anything(), undefined, []);
        });
    });

    describe("#toString()", () => {
        it("should return a non-empty string", async () => {
            const result = await Packer.toString(file);

            assert.isAbove(result.length, 0);
        });
    });

    describe("#toBuffer()", () => {
        it("should create a standard docx file", { timeout: 99_999_999 }, async () => {
            const buffer = await Packer.toBuffer(file);

            assert.isDefined(buffer);
            assert.isTrue(buffer.byteLength > 0);
        });

        it("should handle exception if it throws any", () => {
            vi.spyOn((Packer as any).compiler, "compile").mockImplementation(() => {
                throw new Error();
            });

            return Packer.toBuffer(file).catch((error) => {
                assert.isDefined(error);
            });
        });

        afterEach(() => {
            vi.restoreAllMocks();
        });
    });

    describe("#toBase64String()", () => {
        it("should create a standard docx file", { timeout: 99_999_999 }, async () => {
            const str = await Packer.toBase64String(file);
            expect(str).toBeDefined();
            expect(str.length).toBeGreaterThan(0);
        });

        it("should handle exception if it throws any", () => {
            vi.spyOn((Packer as any).compiler, "compile").mockImplementation(() => {
                throw new Error();
            });

            return Packer.toBase64String(file).catch((error) => {
                assert.isDefined(error);
            });
        });

        afterEach(() => {
            vi.resetAllMocks();
        });
    });

    describe("#toBlob()", () => {
        it("should create a standard docx file", async () => {
            vi.spyOn((Packer as any).compiler, "compile").mockReturnValue({});
            const str = await Packer.toBlob(file);

            assert.isDefined(str);
        });

        it("should handle exception if it throws any", () => {
            vi.spyOn((Packer as any).compiler, "compile").mockImplementation(() => {
                throw new Error();
            });

            return Packer.toBlob(file).catch((error) => {
                assert.isDefined(error);
            });
        });

        afterEach(() => {
            vi.resetAllMocks();
        });
    });

    describe("#toArrayBuffer()", () => {
        it("should create a standard docx file", async () => {
            vi.spyOn((Packer as any).compiler, "compile").mockReturnValue({});
            const str = await Packer.toArrayBuffer(file);

            assert.isDefined(str);
        });

        it("should handle exception if it throws any", () => {
            vi.spyOn((Packer as any).compiler, "compile").mockImplementation(() => {
                throw new Error();
            });

            return Packer.toArrayBuffer(file).catch((error) => {
                assert.isDefined(error);
            });
        });

        afterEach(() => {
            vi.resetAllMocks();
        });
    });

    describe("output types", () => {
        it("should export to uint8array", async () => {
            vi.spyOn((Packer as any).compiler, "compile").mockReturnValue({});

            const result = await Packer.pack(file, "uint8array");
            expect(result).toBeInstanceOf(Uint8Array);
        });

        it("should export to binarystring", async () => {
            vi.spyOn((Packer as any).compiler, "compile").mockReturnValue({});

            const result = await Packer.pack(file, "binarystring");
            expect(typeof result).toBe("string");
        });

        it("should export to array", async () => {
            vi.spyOn((Packer as any).compiler, "compile").mockReturnValue({});

            const result = await Packer.pack(file, "array");
            expect(Array.isArray(result)).toBe(true);
        });

        afterEach(() => {
            vi.restoreAllMocks();
        });
    });
});

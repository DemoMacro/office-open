import { File } from "@file/file";
import { Footer, Header } from "@file/header";
import { ImageRun, Paragraph } from "@file/paragraph";
import * as convenienceFunctions from "@util/convenience-functions";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { Compiler } from "./next-compiler";

describe("Compiler", () => {
    let compiler: Compiler;

    beforeEach(() => {
        compiler = new Compiler();
    });

    beforeAll(() => {
        vi.spyOn(convenienceFunctions, "uniqueId").mockReturnValue("test");
    });

    afterAll(() => {
        vi.resetAllMocks();
    });

    describe("#compile()", () => {
        it("should pack all the content", { timeout: 99_999_999 }, () => {
            const file = new File({
                comments: {
                    children: [],
                },
                sections: [],
            });
            const zipFile = compiler.compile(file);
            const fileNames = Object.keys(zipFile);

            expect(fileNames).is.an.instanceof(Array);
            expect(fileNames).has.length(18);
            expect(fileNames).to.include("word/document.xml");
            expect(fileNames).to.include("word/styles.xml");
            expect(fileNames).to.include("docProps/core.xml");
            expect(fileNames).to.include("docProps/custom.xml");
            expect(fileNames).to.include("docProps/app.xml");
            expect(fileNames).to.include("word/numbering.xml");
            expect(fileNames).to.include("word/footnotes.xml");
            expect(fileNames).to.include("word/_rels/footnotes.xml.rels");
            expect(fileNames).to.include("word/endnotes.xml");
            expect(fileNames).to.include("word/_rels/endnotes.xml.rels");
            expect(fileNames).to.include("word/settings.xml");
            expect(fileNames).to.include("word/comments.xml");
            expect(fileNames).to.include("word/fontTable.xml");
            expect(fileNames).to.include("word/_rels/document.xml.rels");
            expect(fileNames).to.include("word/_rels/fontTable.xml.rels");
            expect(fileNames).to.include("[Content_Types].xml");
            expect(fileNames).to.include("_rels/.rels");
        });

        it("should pack all additional headers and footers", { timeout: 99_999_999 }, () => {
            const file = new File({
                sections: [
                    {
                        children: [],
                        footers: {
                            default: new Footer({
                                children: [new Paragraph("test")],
                            }),
                        },
                        headers: {
                            default: new Header({
                                children: [new Paragraph("test")],
                            }),
                        },
                    },
                    {
                        children: [],
                        footers: {
                            default: new Footer({
                                children: [new Paragraph("test")],
                            }),
                        },
                        headers: {
                            default: new Header({
                                children: [new Paragraph("test")],
                            }),
                        },
                    },
                ],
            });

            const zipFile = compiler.compile(file);
            const fileNames = Object.keys(zipFile);

            expect(fileNames).is.an.instanceof(Array);
            expect(fileNames).has.length(26);

            expect(fileNames).to.include("word/header1.xml");
            expect(fileNames).to.include("word/_rels/header1.xml.rels");
            expect(fileNames).to.include("word/header2.xml");
            expect(fileNames).to.include("word/_rels/header2.xml.rels");
            expect(fileNames).to.include("word/footer1.xml");
            expect(fileNames).to.include("word/_rels/footer1.xml.rels");
            expect(fileNames).to.include("word/footer2.xml");
            expect(fileNames).to.include("word/_rels/footer2.xml.rels");
        });

        it("should pack subfile overrides", { timeout: 99_999_999 }, () => {
            const file = new File({
                comments: {
                    children: [],
                },
                sections: [],
            });
            const subfileData1 = "comments";
            const subfileData2 = "commentsExtended";
            const overrides = [
                { data: subfileData1, path: "word/comments.xml" },
                { data: subfileData2, path: "word/commentsExtended.xml" },
            ];
            const zipFile = compiler.compile(file, "", overrides);
            const fileNames = Object.keys(zipFile);

            expect(fileNames).is.an.instanceof(Array);
            expect(fileNames).has.length(19);

            expect(fileNames).to.include("word/comments.xml");
            expect(fileNames).to.include("word/commentsExtended.xml");
        });

        it("should call the format method X times equalling X files to be formatted", () => {
            // This test is required because before, there was a case where Document was formatted twice, which was inefficient
            // This also caused issues such as running prepForXml multiple times as format() was ran multiple times.
            const paragraph = new Paragraph("");
            const file = new File({
                sections: [
                    {
                        children: [paragraph],
                        properties: {},
                    },
                ],
            });

            const spy = vi.spyOn(compiler["formatter"], "format");

            compiler.compile(file);
            expect(spy).toBeCalledTimes(18);
        });

        it("should work with media datas", () => {
            const file = new File({
                sections: [
                    {
                        children: [
                            new Paragraph({
                                children: [
                                    new ImageRun({
                                        data: Buffer.from("", "base64"),
                                        transformation: {
                                            height: 100,
                                            width: 100,
                                        },
                                        type: "png",
                                    }),
                                    new ImageRun({
                                        data: Buffer.from("", "base64"),
                                        fallback: {
                                            data: Buffer.from("", "base64"),
                                            type: "png",
                                        },
                                        transformation: {
                                            height: 100,
                                            width: 100,
                                        },
                                        type: "svg",
                                    }),
                                ],
                            }),
                        ],
                        footers: {
                            default: new Footer({
                                children: [new Paragraph("test")],
                            }),
                        },
                        headers: {
                            default: new Header({
                                children: [new Paragraph("test")],
                            }),
                        },
                    },
                ],
            });

            vi.spyOn(compiler["imageReplacer"], "getMediaData").mockReturnValue([
                {
                    data: Buffer.from(""),
                    fileName: "test",
                    transformation: {
                        emus: {
                            x: 100,
                            y: 100,
                        },
                        pixels: {
                            x: 100,
                            y: 100,
                        },
                    },
                    type: "png",
                },
                {
                    data: Buffer.from(""),
                    fallback: {
                        data: Buffer.from(""),
                        fileName: "test",
                        transformation: {
                            emus: {
                                x: 100,
                                y: 100,
                            },
                            pixels: {
                                x: 100,
                                y: 100,
                            },
                        },
                        type: "png",
                    },
                    fileName: "test",
                    transformation: {
                        emus: {
                            x: 100,
                            y: 100,
                        },
                        pixels: {
                            x: 100,
                            y: 100,
                        },
                    },
                    type: "svg",
                },
            ]);

            compiler.compile(file);
        });

        it("should work with fonts", () => {
            const file = new File({
                fonts: [{ data: Buffer.from(""), name: "Pacifico" }],
                sections: [],
            });

            compiler.compile(file);
        });
    });
});

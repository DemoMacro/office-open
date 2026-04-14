import { Formatter } from "@export/formatter";
import type { IViewWrapper } from "@file/document-wrapper";
import type { File } from "@file/file";
import { Paragraph } from "@file/index";
import type { IMediaData } from "@file/media";
import { describe, expect, it, vi } from "vite-plus/test";

import { WpgGroupRun } from "./wpg-group-run";

describe("WpgGroupRun", () => {
    describe("#constructor()", () => {
        it("should create a WpgGroupRun with image children", () => {
            const imageChild: IMediaData = {
                data: Buffer.from(""),
                fileName: "test-image.png",
                transformation: {
                    emus: { x: 952_500, y: 952_500 },
                    pixels: { x: 100, y: 100 },
                },
                type: "png",
            };

            const run = new WpgGroupRun({
                children: [imageChild],
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            const tree = new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");
        });

        it("should create a WpgGroupRun with wps shape children", () => {
            const run = new WpgGroupRun({
                children: [
                    {
                        data: {
                            children: [new Paragraph("Shape text")],
                        },
                        transformation: {
                            emus: { x: 952_500, y: 952_500 },
                            pixels: { x: 100, y: 100 },
                        },
                        type: "wps",
                    },
                ],
                transformation: {
                    height: 300,
                    width: 300,
                },
                type: "wpg",
            });

            const tree = new Formatter().format(run, {
                file: {
                    Media: {},
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");
        });

        it("should create a WpgGroupRun with mixed children", () => {
            const imageChild: IMediaData = {
                data: Buffer.from(""),
                fileName: "mixed-image.png",
                transformation: {
                    emus: { x: 1_428_750, y: 1_428_750 },
                    pixels: { x: 150, y: 150 },
                },
                type: "png",
            };

            const run = new WpgGroupRun({
                children: [
                    imageChild,
                    {
                        data: {
                            children: [new Paragraph("Text in shape")],
                        },
                        transformation: {
                            emus: { x: 952_500, y: 952_500 },
                            pixels: { x: 100, y: 100 },
                        },
                        type: "wps",
                    },
                ],
                transformation: {
                    height: 400,
                    width: 400,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            const tree = new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");
        });

        it("should support floating positioning", () => {
            const run = new WpgGroupRun({
                children: [],
                floating: {
                    horizontalPosition: {
                        offset: 1_014_400,
                    },
                    verticalPosition: {
                        offset: 1_014_400,
                    },
                    zIndex: 5,
                },
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const tree = new Formatter().format(run, {
                file: {
                    Media: {},
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");

            const drawing = (tree as Record<string, readonly unknown[]>)["w:r"];
            const drawingElement = drawing.find(
                (el) => typeof el === "object" && el !== null && "w:drawing" in el,
            );
            expect(drawingElement).toBeDefined();
        });

        it("should support transformation with offset and rotation", () => {
            const run = new WpgGroupRun({
                children: [],
                transformation: {
                    height: 100,
                    offset: { left: 10, top: 20 },
                    rotation: 90,
                    width: 200,
                },
                type: "wpg",
            });

            const tree = new Formatter().format(run, {
                file: {
                    Media: {},
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");
        });

        it("should support altText option", () => {
            const run = new WpgGroupRun({
                altText: {
                    description: "Group Description",
                    name: "Group Name",
                    title: "Group Title",
                },
                children: [],
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const tree = new Formatter().format(run, {
                file: {
                    Media: {},
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).toBeDefined();
            expect(tree).toHaveProperty("w:r");
        });
    });

    describe("#prepForXml()", () => {
        it("should add image children to context.file.Media", () => {
            const imageChild: IMediaData = {
                data: Buffer.from("image-data"),
                fileName: "test.png",
                transformation: {
                    emus: { x: 952_500, y: 952_500 },
                    pixels: { x: 100, y: 100 },
                },
                type: "png",
            };

            const run = new WpgGroupRun({
                children: [imageChild],
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(addImageMock).toHaveBeenCalledWith("test.png", imageChild);
        });

        it("should not call addImage for wps children", () => {
            const run = new WpgGroupRun({
                children: [
                    {
                        data: {
                            children: [new Paragraph("Test")],
                        },
                        transformation: {
                            emus: { x: 952_500, y: 952_500 },
                            pixels: { x: 100, y: 100 },
                        },
                        type: "wps",
                    },
                ],
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(addImageMock).not.toHaveBeenCalled();
        });

        it("should add SVG fallback image to media", () => {
            const svgChild: IMediaData = {
                data: Buffer.from("svg-data"),
                fallback: {
                    data: Buffer.from("fallback-data"),
                    fileName: "test-fallback.png",
                    transformation: {
                        emus: { x: 952_500, y: 952_500 },
                        pixels: { x: 100, y: 100 },
                    },
                    type: "png",
                },
                fileName: "test.svg",
                transformation: {
                    emus: { x: 952_500, y: 952_500 },
                    pixels: { x: 100, y: 100 },
                },
                type: "svg",
            };

            const run = new WpgGroupRun({
                children: [svgChild],
                transformation: {
                    height: 200,
                    width: 200,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(addImageMock).toHaveBeenCalledTimes(2);
            expect(addImageMock).toHaveBeenCalledWith("test.svg", svgChild);
            expect(addImageMock).toHaveBeenCalledWith("test-fallback.png", svgChild.fallback);
        });

        it("should add multiple image children to media", () => {
            const imageChild1: IMediaData = {
                data: Buffer.from("data1"),
                fileName: "image1.png",
                transformation: {
                    emus: { x: 952_500, y: 952_500 },
                    pixels: { x: 100, y: 100 },
                },
                type: "png",
            };

            const imageChild2: IMediaData = {
                data: Buffer.from("data2"),
                fileName: "image2.jpg",
                transformation: {
                    emus: { x: 1_905_000, y: 1_905_000 },
                    pixels: { x: 200, y: 200 },
                },
                type: "jpg",
            };

            const run = new WpgGroupRun({
                children: [imageChild1, imageChild2],
                transformation: {
                    height: 400,
                    width: 400,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(addImageMock).toHaveBeenCalledTimes(2);
            expect(addImageMock).toHaveBeenCalledWith("image1.png", imageChild1);
            expect(addImageMock).toHaveBeenCalledWith("image2.jpg", imageChild2);
        });

        it("should only add non-wps children to media when mixed", () => {
            const imageChild: IMediaData = {
                data: Buffer.from("data"),
                fileName: "image.png",
                transformation: {
                    emus: { x: 952_500, y: 952_500 },
                    pixels: { x: 100, y: 100 },
                },
                type: "png",
            };

            const run = new WpgGroupRun({
                children: [
                    imageChild,
                    {
                        data: {
                            children: [new Paragraph("Text")],
                        },
                        transformation: {
                            emus: { x: 952_500, y: 952_500 },
                            pixels: { x: 100, y: 100 },
                        },
                        type: "wps",
                    },
                ],
                transformation: {
                    height: 400,
                    width: 400,
                },
                type: "wpg",
            });

            const addImageMock = vi.fn();
            new Formatter().format(run, {
                file: {
                    Media: {
                        addImage: addImageMock,
                    },
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(addImageMock).toHaveBeenCalledTimes(1);
            expect(addImageMock).toHaveBeenCalledWith("image.png", imageChild);
        });
    });
});

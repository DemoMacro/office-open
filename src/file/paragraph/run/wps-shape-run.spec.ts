import { Formatter } from "@export/formatter";
import type { IViewWrapper } from "@file/document-wrapper";
import type { File } from "@file/file";
import { Paragraph } from "@file/index";
import { describe, expect, it } from "vite-plus/test";

import { WpsShapeRun } from "./wps-shape-run";

describe("WpsShapeRun", () => {
    describe("#constructor()", () => {
        it("should create with Buffer", () => {
            const currentShapeRun = new WpsShapeRun({
                children: [new Paragraph("Test Paragraph")],
                floating: {
                    horizontalPosition: {
                        offset: 1_014_400,
                    },
                    verticalPosition: {
                        offset: 1_014_400,
                    },
                    zIndex: 10,
                },
                solidFill: {
                    value: "FF0000",
                },
                transformation: {
                    height: 200,
                    rotation: 45,
                    width: 200,
                },
                type: "wps",
            });

            const tree = new Formatter().format(currentShapeRun, {
                file: {
                    Media: {},
                } as unknown as File,
                stack: [],
                viewWrapper: {} as unknown as IViewWrapper,
            });

            expect(tree).to.deep.equal({
                "w:r": [
                    {
                        "w:drawing": [
                            {
                                "wp:anchor": [
                                    {
                                        _attr: {
                                            distT: 0,
                                            distB: 0,
                                            distL: 0,
                                            distR: 0,
                                            simplePos: "0",
                                            allowOverlap: "1",
                                            behindDoc: "0",
                                            locked: "0",
                                            layoutInCell: "1",
                                            relativeHeight: 10,
                                        },
                                    },
                                    {
                                        "wp:simplePos": {
                                            _attr: {
                                                x: 0,
                                                y: 0,
                                            },
                                        },
                                    },
                                    {
                                        "wp:positionH": [
                                            {
                                                _attr: {
                                                    relativeFrom: "page",
                                                },
                                            },
                                            {
                                                "wp:posOffset": ["1014400"],
                                            },
                                        ],
                                    },
                                    {
                                        "wp:positionV": [
                                            {
                                                _attr: {
                                                    relativeFrom: "page",
                                                },
                                            },
                                            {
                                                "wp:posOffset": ["1014400"],
                                            },
                                        ],
                                    },
                                    {
                                        "wp:extent": {
                                            _attr: {
                                                cx: 1905000,
                                                cy: 1905000,
                                            },
                                        },
                                    },
                                    {
                                        "wp:effectExtent": {
                                            _attr: {
                                                b: 0,
                                                l: 0,
                                                r: 0,
                                                t: 0,
                                            },
                                        },
                                    },
                                    {
                                        "wp:wrapNone": {},
                                    },
                                    {
                                        "wp:docPr": {
                                            _attr: {
                                                id: 1,
                                                name: "",
                                                descr: "",
                                                title: "",
                                            },
                                        },
                                    },
                                    {
                                        "wp:cNvGraphicFramePr": [
                                            {
                                                "a:graphicFrameLocks": {
                                                    _attr: {
                                                        noChangeAspect: 1,
                                                        "xmlns:a":
                                                            "http://schemas.openxmlformats.org/drawingml/2006/main",
                                                    },
                                                },
                                            },
                                        ],
                                    },
                                    {
                                        "a:graphic": [
                                            {
                                                _attr: {
                                                    "xmlns:a":
                                                        "http://schemas.openxmlformats.org/drawingml/2006/main",
                                                },
                                            },
                                            {
                                                "a:graphicData": [
                                                    {
                                                        _attr: {
                                                            uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                                                        },
                                                    },
                                                    {
                                                        "wps:wsp": [
                                                            {
                                                                "wps:cNvSpPr": {
                                                                    _attr: {
                                                                        txBox: "1",
                                                                    },
                                                                },
                                                            },
                                                            {
                                                                "wps:spPr": [
                                                                    {
                                                                        _attr: {
                                                                            bwMode: "auto",
                                                                        },
                                                                    },
                                                                    {
                                                                        "a:xfrm": [
                                                                            {
                                                                                _attr: {
                                                                                    rot: 2700000,
                                                                                },
                                                                            },
                                                                            {
                                                                                "a:off": {
                                                                                    _attr: {
                                                                                        x: 0,
                                                                                        y: 0,
                                                                                    },
                                                                                },
                                                                            },
                                                                            {
                                                                                "a:ext": {
                                                                                    _attr: {
                                                                                        cx: 1905000,
                                                                                        cy: 1905000,
                                                                                    },
                                                                                },
                                                                            },
                                                                        ],
                                                                    },
                                                                    {
                                                                        "a:prstGeom": [
                                                                            {
                                                                                _attr: {
                                                                                    prst: "rect",
                                                                                },
                                                                            },
                                                                            {
                                                                                "a:avLst": {},
                                                                            },
                                                                        ],
                                                                    },
                                                                    {
                                                                        "a:solidFill": [
                                                                            {
                                                                                "a:srgbClr": {
                                                                                    _attr: {
                                                                                        val: "FF0000",
                                                                                    },
                                                                                },
                                                                            },
                                                                        ],
                                                                    },
                                                                ],
                                                            },
                                                            {
                                                                "wps:txbx": [
                                                                    {
                                                                        "w:txbxContent": [
                                                                            {
                                                                                "w:p": [
                                                                                    {
                                                                                        "w:r": [
                                                                                            {
                                                                                                "w:t": [
                                                                                                    {
                                                                                                        _attr: {
                                                                                                            "xml:space":
                                                                                                                "preserve",
                                                                                                        },
                                                                                                    },
                                                                                                    "Test Paragraph",
                                                                                                ],
                                                                                            },
                                                                                        ],
                                                                                    },
                                                                                ],
                                                                            },
                                                                        ],
                                                                    },
                                                                ],
                                                            },
                                                            {
                                                                "wps:bodyPr": {},
                                                            },
                                                        ],
                                                    },
                                                ],
                                            },
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });
    });
});

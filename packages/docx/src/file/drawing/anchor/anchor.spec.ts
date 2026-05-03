import { Formatter } from "@export/formatter";
import { Utility } from "tests/utility";
import { assert, beforeEach, describe, expect, it } from "vite-plus/test";

import { resetDocPropertiesIdGen } from "../doc-properties/doc-properties";
import type { IDrawingOptions } from "../drawing";
import { TextWrappingSide, TextWrappingType } from "../text-wrap";
import { Anchor } from "./anchor";

const createAnchor = (drawingOptions: IDrawingOptions): Anchor =>
    new Anchor({
        drawingOptions,
        mediaData: {
            data: Buffer.from(""),
            fileName: "test.png",
            transformation: {
                emus: {
                    x: 0,
                    y: 0,
                },
                pixels: {
                    x: 0,
                    y: 0,
                },
            },
            type: "png",
        },
        transform: {
            emus: {
                x: 100 * 9525,
                y: 100 * 9525,
            },
            pixels: {
                x: 100,
                y: 100,
            },
        },
    });

describe("Anchor", () => {
    let anchor: Anchor;

    beforeEach(() => {
        resetDocPropertiesIdGen();
    });

    describe("#constructor()", () => {
        it("should create a Drawing with correct root key", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.rootKey, "wp:anchor");
            assert.equal(newJson.root.length, 10);
        });

        it("should create a Drawing with all default options", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                allowOverlap: "1",
                behindDoc: "0",
                distB: 0,
                distL: 0,
                distR: 0,
                distT: 0,
                layoutInCell: "1",
                locked: "0",
                relativeHeight: 952_500,
                simplePos: "0",
            });

            // 1: simple pos
            assert.equal(newJson.root[1].rootKey, "wp:simplePos");

            // 2: horizontal position
            const horizontalPosition = newJson.root[2];
            assert.equal(horizontalPosition.rootKey, "wp:positionH");

            assert.equal(horizontalPosition.root[1].rootKey, "wp:posOffset");
            assert.include(horizontalPosition.root[1].root[0], 0);

            // 3: vertical position
            const verticalPosition = newJson.root[3];
            assert.equal(verticalPosition.rootKey, "wp:positionV");

            assert.equal(verticalPosition.root[1].rootKey, "wp:posOffset");
            assert.include(verticalPosition.root[1].root[0], 0);

            // 4: extent
            const extent = newJson.root[4];
            assert.equal(extent.rootKey, "wp:extent");

            // 5: effect extent
            const effectExtent = newJson.root[5];
            assert.equal(effectExtent.rootKey, "wp:effectExtent");

            // 6 text wrap: none
            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapNone");

            // 7: doc properties
            const docProperties = newJson.root[7];
            assert.equal(docProperties.rootKey, "wp:docPr");

            // 8: graphic frame properties
            const graphicFrame = newJson.root[8];
            assert.equal(graphicFrame.rootKey, "wp:cNvGraphicFramePr");

            // 9: graphic
            const graphic = newJson.root[9];
            assert.equal(graphic.rootKey, "a:graphic");
        });

        it("should create a Drawing with square text wrapping", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    wrap: {
                        type: TextWrappingType.SQUARE,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            // 6 text wrap: square
            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapSquare");
        });

        it("should create a Drawing with no text wrapping", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    wrap: {
                        type: TextWrappingType.NONE,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapNone");
        });

        it("should create a Drawing with tight text wrapping", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    wrap: {
                        type: TextWrappingType.TIGHT,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapTight");
        });

        it("should create a Drawing with tight text wrapping", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    wrap: {
                        type: TextWrappingType.TOP_AND_BOTTOM,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapTopAndBottom");
        });

        it("should create a Drawing with through text wrapping", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    wrap: {
                        side: TextWrappingSide.BOTH_SIDES,
                        type: TextWrappingType.THROUGH,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            assert.equal(newJson.root.length, 10);

            const textWrap = newJson.root[6];
            assert.equal(textWrap.rootKey, "wp:wrapThrough");
        });

        it("should create a Drawing with a margin", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    margins: {
                        bottom: 10,
                        left: 10,
                        right: 10,
                        top: 10,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                distB: 10,
                distL: 10,
                distR: 10,
                distT: 10,
            });
        });

        it("should create a Drawing with a default margin", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    margins: {},
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                distB: 0,
                distL: 0,
                distR: 0,
                distT: 0,
            });
        });

        it("should create a Drawing with allowOverlap being false", () => {
            anchor = createAnchor({
                floating: {
                    allowOverlap: false,
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                allowOverlap: "0",
            });
        });

        it("should create a Drawing with behindDocument being true", () => {
            anchor = createAnchor({
                floating: {
                    behindDocument: true,
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                behindDoc: "1",
            });
        });

        it("should create a Drawing with locked being true", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    lockAnchor: true,
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                locked: "1",
            });
        });

        it("should create a Drawing with locked being false", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    layoutInCell: false,
                    verticalPosition: {
                        offset: 0,
                    },
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;
            assert.include(anchorAttributes, {
                layoutInCell: "0",
            });
        });

        it("should create a Drawing with a certain z-index", () => {
            anchor = createAnchor({
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    zIndex: 120,
                },
            });
            const newJson = Utility.jsonify(anchor);
            const anchorAttributes = newJson.root[0].root;

            assert.include(anchorAttributes, {
                relativeHeight: 120,
            });
        });

        it("should create a Drawing with doc properties", () => {
            anchor = createAnchor({
                docProperties: {
                    description: "test",
                    name: "test",
                    title: "test",
                },
                floating: {
                    horizontalPosition: {
                        offset: 0,
                    },
                    verticalPosition: {
                        offset: 0,
                    },
                    zIndex: 120,
                },
            });
            const tree = new Formatter().format(anchor);
            expect(tree).to.deep.equal({
                "wp:anchor": [
                    {
                        _attr: {
                            allowOverlap: "1",
                            behindDoc: "0",
                            distB: 0,
                            distL: 0,
                            distR: 0,
                            distT: 0,
                            layoutInCell: "1",
                            locked: "0",
                            relativeHeight: 120,
                            simplePos: "0",
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
                                "wp:posOffset": ["0"],
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
                                "wp:posOffset": ["0"],
                            },
                        ],
                    },
                    {
                        "wp:extent": {
                            _attr: {
                                cx: 952_500,
                                cy: 952_500,
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
                                descr: "test",
                                id: 1,
                                name: "test",
                                title: "test",
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
                                            uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
                                        },
                                    },
                                    {
                                        "pic:pic": [
                                            {
                                                _attr: {
                                                    "xmlns:pic":
                                                        "http://schemas.openxmlformats.org/drawingml/2006/picture",
                                                },
                                            },
                                            {
                                                "pic:nvPicPr": [
                                                    {
                                                        "pic:cNvPr": {
                                                            _attr: {
                                                                descr: "",
                                                                id: 0,
                                                                name: "",
                                                            },
                                                        },
                                                    },
                                                    {
                                                        "pic:cNvPicPr": [
                                                            {
                                                                _attr: {
                                                                    preferRelativeResize: true,
                                                                },
                                                            },
                                                            {
                                                                "a:picLocks": {
                                                                    _attr: {
                                                                        noChangeArrowheads: 1,
                                                                        noChangeAspect: 1,
                                                                    },
                                                                },
                                                            },
                                                        ],
                                                    },
                                                ],
                                            },
                                            {
                                                "pic:blipFill": [
                                                    {
                                                        "a:blip": {
                                                            _attr: {
                                                                cstate: "none",
                                                                "r:embed": "{test.png}",
                                                            },
                                                        },
                                                    },
                                                    {
                                                        "a:srcRect": {},
                                                    },
                                                    {
                                                        "a:stretch": [
                                                            {
                                                                "a:fillRect": {},
                                                            },
                                                        ],
                                                    },
                                                ],
                                            },
                                            {
                                                "pic:spPr": [
                                                    {
                                                        _attr: {
                                                            bwMode: "auto",
                                                        },
                                                    },
                                                    {
                                                        "a:xfrm": [
                                                            {
                                                                _attr: {},
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
                                                                        cx: 952_500,
                                                                        cy: 952_500,
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

import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createCustomGeometry } from "./custom-geometry";

describe("createCustomGeometry", () => {
    it("should create a custom geometry with a single path", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "0" } },
                            { command: "lineTo", point: { x: "10000000", y: "0" } },
                            { command: "lineTo", point: { x: "5000000", y: "10000000" } },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "0" } } }] },
                                { "a:lnTo": [{ "a:pt": { _attr: { x: "10000000", y: "0" } } }] },
                                {
                                    "a:lnTo": [
                                        { "a:pt": { _attr: { x: "5000000", y: "10000000" } } },
                                    ],
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with path attributes", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        w: 10000000,
                        h: 10000000,
                        fill: "norm",
                        stroke: true,
                        extrusionOk: false,
                        commands: [{ command: "close" }],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                {
                                    _attr: {
                                        w: 10000000,
                                        h: 10000000,
                                        fill: "norm",
                                        stroke: true,
                                        extrusionOk: false,
                                    },
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with arcTo command", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "0" } },
                            {
                                command: "arcTo",
                                wR: "5000000",
                                hR: "5000000",
                                stAng: "0",
                                swAng: "5400000",
                            },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "0" } } }] },
                                {
                                    "a:arcTo": {
                                        _attr: {
                                            wR: "5000000",
                                            hR: "5000000",
                                            stAng: "0",
                                            swAng: "5400000",
                                        },
                                    },
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with quadBezTo command", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "0" } },
                            {
                                command: "quadBezTo",
                                points: [
                                    { x: "5000000", y: "-2000000" },
                                    { x: "10000000", y: "0" },
                                ],
                            },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "0" } } }] },
                                {
                                    "a:quadBezTo": [
                                        { "a:pt": { _attr: { x: "5000000", y: "-2000000" } } },
                                        { "a:pt": { _attr: { x: "10000000", y: "0" } } },
                                    ],
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with cubicBezTo command", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "0" } },
                            {
                                command: "cubicBezTo",
                                points: [
                                    { x: "2000000", y: "-2000000" },
                                    { x: "8000000", y: "-2000000" },
                                    { x: "10000000", y: "0" },
                                ],
                            },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "0" } } }] },
                                {
                                    "a:cubicBezTo": [
                                        { "a:pt": { _attr: { x: "2000000", y: "-2000000" } } },
                                        { "a:pt": { _attr: { x: "8000000", y: "-2000000" } } },
                                        { "a:pt": { _attr: { x: "10000000", y: "0" } } },
                                    ],
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with adjustment values and guides", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                adjustmentValues: [{ name: "adj", formula: "val 50000" }],
                guides: [{ name: "myGuide", formula: "*/ 10000000 2" }],
                pathList: [{ commands: [{ command: "close" }] }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:avLst": [{ "a:gd": { _attr: { name: "adj", fmla: "val 50000" } } }],
                },
                {
                    "a:gdLst": [{ "a:gd": { _attr: { name: "myGuide", fmla: "*/ 10000000 2" } } }],
                },
                {
                    "a:pathLst": [{ "a:path": [{ "a:close": {} }] }],
                },
            ],
        });
    });

    it("should create a custom geometry with adjust handles", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                adjustHandles: [
                    {
                        type: "xy",
                        gdRefX: "adjX",
                        minX: "0",
                        maxX: "10000000",
                        position: { x: "5000000", y: "5000000" },
                    },
                    {
                        type: "polar",
                        gdRefR: "adjR",
                        minR: "0",
                        maxR: "10000000",
                        gdRefAng: "adjAng",
                        position: { x: "5000000", y: "0" },
                    },
                ],
                pathList: [{ commands: [{ command: "close" }] }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:ahLst": [
                        {
                            "a:ahXY": [
                                { _attr: { gdRefX: "adjX", minX: "0", maxX: "10000000" } },
                                { "a:pos": { _attr: { x: "5000000", y: "5000000" } } },
                            ],
                        },
                        {
                            "a:ahPolar": [
                                {
                                    _attr: {
                                        gdRefR: "adjR",
                                        minR: "0",
                                        maxR: "10000000",
                                        gdRefAng: "adjAng",
                                    },
                                },
                                { "a:pos": { _attr: { x: "5000000", y: "0" } } },
                            ],
                        },
                    ],
                },
                {
                    "a:pathLst": [{ "a:path": [{ "a:close": {} }] }],
                },
            ],
        });
    });

    it("should create a custom geometry with connection sites", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                connectionSites: [
                    { ang: "0", position: { x: "5000000", y: "0" } },
                    { ang: "5400000", position: { x: "10000000", y: "5000000" } },
                ],
                pathList: [{ commands: [{ command: "close" }] }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:cxnLst": [
                        {
                            "a:cxn": [
                                { _attr: { ang: "0" } },
                                { "a:pos": { _attr: { x: "5000000", y: "0" } } },
                            ],
                        },
                        {
                            "a:cxn": [
                                { _attr: { ang: "5400000" } },
                                { "a:pos": { _attr: { x: "10000000", y: "5000000" } } },
                            ],
                        },
                    ],
                },
                {
                    "a:pathLst": [{ "a:path": [{ "a:close": {} }] }],
                },
            ],
        });
    });

    it("should create a custom geometry with text rect", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                textRect: { l: "0", t: "0", r: "10000000", b: "10000000" },
                pathList: [{ commands: [{ command: "close" }] }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:rect": { _attr: { l: "0", t: "0", r: "10000000", b: "10000000" } },
                },
                {
                    "a:pathLst": [{ "a:path": [{ "a:close": {} }] }],
                },
            ],
        });
    });

    it("should create a custom geometry with multiple paths", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                pathList: [
                    {
                        w: 10000000,
                        h: 10000000,
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "0" } },
                            { command: "lineTo", point: { x: "10000000", y: "0" } },
                            { command: "close" },
                        ],
                    },
                    {
                        w: 10000000,
                        h: 10000000,
                        commands: [
                            { command: "moveTo", point: { x: "0", y: "10000000" } },
                            { command: "lineTo", point: { x: "10000000", y: "10000000" } },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { _attr: { w: 10000000, h: 10000000 } },
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "0" } } }] },
                                { "a:lnTo": [{ "a:pt": { _attr: { x: "10000000", y: "0" } } }] },
                                { "a:close": {} },
                            ],
                        },
                        {
                            "a:path": [
                                { _attr: { w: 10000000, h: 10000000 } },
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "0", y: "10000000" } } }] },
                                {
                                    "a:lnTo": [
                                        { "a:pt": { _attr: { x: "10000000", y: "10000000" } } },
                                    ],
                                },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a custom geometry with all optional sections", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                adjustmentValues: [{ name: "adj", formula: "val 50000" }],
                guides: [{ name: "g1", formula: "val 100" }],
                adjustHandles: [{ type: "xy", position: { x: "0", y: "0" } }],
                connectionSites: [{ ang: "0", position: { x: "0", y: "0" } }],
                textRect: { l: "l", t: "t", r: "r", b: "b" },
                pathList: [{ commands: [{ command: "close" }] }],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                { "a:avLst": [{ "a:gd": { _attr: { name: "adj", fmla: "val 50000" } } }] },
                { "a:gdLst": [{ "a:gd": { _attr: { name: "g1", fmla: "val 100" } } }] },
                {
                    "a:ahLst": [
                        {
                            "a:ahXY": [{ "a:pos": { _attr: { x: "0", y: "0" } } }],
                        },
                    ],
                },
                {
                    "a:cxnLst": [
                        {
                            "a:cxn": [
                                { _attr: { ang: "0" } },
                                { "a:pos": { _attr: { x: "0", y: "0" } } },
                            ],
                        },
                    ],
                },
                { "a:rect": { _attr: { l: "l", t: "t", r: "r", b: "b" } } },
                { "a:pathLst": [{ "a:path": [{ "a:close": {} }] }] },
            ],
        });
    });

    it("should use guides as references in path commands", () => {
        const tree = new Formatter().format(
            createCustomGeometry({
                guides: [
                    { name: "w2", formula: "*/ 10000000 2" },
                    { name: "h2", formula: "*/ 10000000 2" },
                ],
                pathList: [
                    {
                        commands: [
                            { command: "moveTo", point: { x: "w2", y: "0" } },
                            { command: "lineTo", point: { x: "10000000", y: "h2" } },
                            { command: "close" },
                        ],
                    },
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "a:custGeom": [
                {
                    "a:gdLst": [
                        { "a:gd": { _attr: { name: "w2", fmla: "*/ 10000000 2" } } },
                        { "a:gd": { _attr: { name: "h2", fmla: "*/ 10000000 2" } } },
                    ],
                },
                {
                    "a:pathLst": [
                        {
                            "a:path": [
                                { "a:moveTo": [{ "a:pt": { _attr: { x: "w2", y: "0" } } }] },
                                { "a:lnTo": [{ "a:pt": { _attr: { x: "10000000", y: "h2" } } }] },
                                { "a:close": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });
});

import { Formatter } from "@export/formatter";
import { beforeEach, describe, expect, it } from "vite-plus/test";

import { TextRun } from "../run";
import { ConcreteHyperlink, ExternalHyperlink, InternalHyperlink } from "./hyperlink";

describe("ConcreteHyperlink", () => {
    let hyperlink: ConcreteHyperlink;

    beforeEach(() => {
        hyperlink = new ConcreteHyperlink(
            [
                new TextRun({
                    style: "Hyperlink",
                    text: "https://example.com",
                }),
            ],
            "superid",
        );
    });

    describe("#constructor()", () => {
        it("should create a hyperlink with correct root key", () => {
            const tree = new Formatter().format(hyperlink);
            expect(tree).to.deep.equal({
                "w:hyperlink": [
                    {
                        _attr: {
                            "r:id": "rIdsuperid",
                            "w:history": 1,
                        },
                    },
                    {
                        "w:r": [
                            { "w:rPr": [{ "w:rStyle": { _attr: { "w:val": "Hyperlink" } } }] },
                            {
                                "w:t": [
                                    { _attr: { "xml:space": "preserve" } },
                                    "https://example.com",
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        describe("with optional anchor parameter", () => {
            beforeEach(() => {
                hyperlink = new ConcreteHyperlink(
                    [
                        new TextRun({
                            style: "Hyperlink",
                            text: "Anchor Text",
                        }),
                    ],
                    "superid2",
                    "anchor",
                );
            });

            it("should create an internal link with anchor tag", () => {
                const tree = new Formatter().format(hyperlink);
                expect(tree).to.deep.equal({
                    "w:hyperlink": [
                        {
                            _attr: {
                                "w:anchor": "anchor",
                                "w:history": 1,
                            },
                        },
                        {
                            "w:r": [
                                { "w:rPr": [{ "w:rStyle": { _attr: { "w:val": "Hyperlink" } } }] },
                                { "w:t": [{ _attr: { "xml:space": "preserve" } }, "Anchor Text"] },
                            ],
                        },
                    ],
                });
            });
        });
    });
});

describe("ExternalHyperlink", () => {
    describe("#constructor()", () => {
        it("should create", () => {
            const externalHyperlink = new ExternalHyperlink({
                children: [new TextRun("test")],
                link: "http://www.google.com",
            });

            expect(externalHyperlink.options.link).to.equal("http://www.google.com");
        });
    });
});

describe("InternalHyperlink", () => {
    describe("#constructor()", () => {
        it("should create", () => {
            const internalHyperlink = new InternalHyperlink({
                anchor: "test-id",
                children: [new TextRun("test")],
            });

            const tree = new Formatter().format(internalHyperlink);

            expect(tree).to.deep.equal({
                "w:hyperlink": [
                    {
                        _attr: {
                            "w:anchor": "test-id",
                            "w:history": 1,
                        },
                    },
                    {
                        "w:r": [
                            {
                                "w:t": [
                                    {
                                        _attr: {
                                            "xml:space": "preserve",
                                        },
                                    },
                                    "test",
                                ],
                            },
                        ],
                    },
                ],
            });
        });
    });
});

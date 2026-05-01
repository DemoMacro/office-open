import { Formatter } from "@export/formatter";
import { StructuredDocumentTagProperties } from "@file/table-of-contents";
import { BuilderElement } from "@file/xml-components";
import { describe, expect, it } from "vite-plus/test";

import { StructuredDocumentTagRun } from "./sdt";

describe("StructuredDocumentTagRun", () => {
    it("should create an inline SDT with text type", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: { text: {}, alias: "Name" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        { "w:alias": { _attr: { "w:val": "Name" } } },
                        { "w:text": { _attr: { "w:multiLine": false } } },
                    ],
                },
            ],
        });
    });

    it("should create an inline SDT with text multiLine", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: { text: { multiLine: true } },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": true } } }],
                },
            ],
        });
    });

    it("should create an inline SDT with comboBox", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    alias: "Color",
                    comboBox: {
                        items: [
                            { displayText: "Red", value: "r" },
                            { displayText: "Blue", value: "b" },
                        ],
                        lastValue: "Red",
                    },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        { "w:alias": { _attr: { "w:val": "Color" } } },
                        {
                            "w:comboBox": [
                                { _attr: { "w:lastValue": "Red" } },
                                {
                                    "w:listItem": {
                                        _attr: { "w:displayText": "Red", "w:value": "r" },
                                    },
                                },
                                {
                                    "w:listItem": {
                                        _attr: { "w:displayText": "Blue", "w:value": "b" },
                                    },
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an inline SDT with dropDownList", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    dropDownList: {
                        items: [{ displayText: "Yes" }, { displayText: "No" }],
                    },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        {
                            "w:dropDownList": [
                                {
                                    "w:listItem": {
                                        _attr: { "w:displayText": "Yes", "w:value": "Yes" },
                                    },
                                },
                                {
                                    "w:listItem": {
                                        _attr: { "w:displayText": "No", "w:value": "No" },
                                    },
                                },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an inline SDT with date type", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    alias: "PickDate",
                    date: {
                        dateFormat: "yyyy-MM-dd",
                        languageId: "en-US",
                        fullDate: "2026-01-15T00:00:00Z",
                    },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        { "w:alias": { _attr: { "w:val": "PickDate" } } },
                        {
                            "w:date": [
                                { _attr: { "w:fullDate": "2026-01-15T00:00:00Z" } },
                                { "w:dateFormat": { _attr: { "w:val": "yyyy-MM-dd" } } },
                                { "w:lid": { _attr: { "w:val": "en-US" } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an SDT with date mapping type", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    date: {
                        dateFormat: "yyyy-MM-dd",
                        storeMappedDataAs: "date",
                        calendar: "gregorian",
                    },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        {
                            "w:date": [
                                { "w:dateFormat": { _attr: { "w:val": "yyyy-MM-dd" } } },
                                { "w:storeMappedDataAs": { _attr: { "w:val": "date" } } },
                                { "w:calendar": { _attr: { "w:val": "gregorian" } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an SDT with common properties", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    alias: "My Control",
                    tag: "myTag",
                    id: 42,
                    lock: "sdtLocked",
                    temporary: true,
                    showingPlaceholder: true,
                    label: 1,
                    tabIndex: 5,
                    text: {},
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        { "w:alias": { _attr: { "w:val": "My Control" } } },
                        { "w:tag": { _attr: { "w:val": "myTag" } } },
                        { "w:id": { _attr: { "w:val": 42 } } },
                        { "w:lock": { _attr: { "w:val": "sdtLocked" } } },
                        { "w:temporary": {} },
                        { "w:showingPlcHdr": {} },
                        { "w:label": { _attr: { "w:val": 1 } } },
                        { "w:tabIndex": { _attr: { "w:val": 5 } } },
                        { "w:text": { _attr: { "w:multiLine": false } } },
                    ],
                },
            ],
        });
    });

    it("should create an SDT with dataBinding", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    dataBinding: {
                        prefixMappings: "xmlns:ns='http://example.com'",
                        xpath: "/ns:root/ns:field",
                        storeItemID: "{00000000-0000-0000-0000-000000000000}",
                    },
                    text: {},
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        {
                            "w:dataBinding": {
                                _attr: {
                                    "w:prefixMappings": "xmlns:ns='http://example.com'",
                                    "w:xpath": "/ns:root/ns:field",
                                    "w:storeItemID": "{00000000-0000-0000-0000-000000000000}",
                                },
                            },
                        },
                        { "w:text": { _attr: { "w:multiLine": false } } },
                    ],
                },
            ],
        });
    });

    it("should create an SDT with empty type (equation, picture, richText, citation, group, bibliography)", () => {
        for (const type of [
            "equation",
            "picture",
            "richText",
            "citation",
            "group",
            "bibliography",
        ] as const) {
            const tree = new Formatter().format(
                new StructuredDocumentTagRun({
                    properties: { [type]: true },
                }),
            );
            expect(tree).to.deep.equal({
                "w:sdt": [
                    {
                        "w:sdtPr": [{ [`w:${type}`]: {} }],
                    },
                ],
            });
        }
    });

    it("should create an SDT with docPartObj", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    docPartObj: { gallery: "Cover Pages", category: "Built-In", unique: true },
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        {
                            "w:docPartObj": [
                                { "w:docPartGallery": { _attr: { "w:val": "Cover Pages" } } },
                                { "w:docPartCategory": { _attr: { "w:val": "Built-In" } } },
                                { "w:docPartUnique": {} },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should create an SDT with children content", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: { text: {} },
                children: [
                    new BuilderElement({
                        name: "w:r",
                        children: [
                            new BuilderElement({
                                name: "w:t",
                            }),
                        ],
                    }),
                ],
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                { "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] },
                {
                    "w:sdtContent": [{ "w:r": [{ "w:t": {} }] }],
                },
            ],
        });
    });

    it("should create an SDT with no type discriminator (empty properties)", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: { alias: "Placeholder" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "Placeholder" } } }],
                },
            ],
        });
    });

    it("should create an SDT with placeholder containing block-level content", () => {
        const tree = new Formatter().format(
            new StructuredDocumentTagRun({
                properties: {
                    alias: "Title",
                    placeholder: [
                        new BuilderElement({
                            name: "w:p",
                            children: [
                                new BuilderElement({
                                    name: "w:r",
                                    children: [
                                        new BuilderElement({
                                            name: "w:t",
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                    text: {},
                },
            }),
        );
        expect(tree).to.deep.equal({
            "w:sdt": [
                {
                    "w:sdtPr": [
                        { "w:alias": { _attr: { "w:val": "Title" } } },
                        { "w:placeholder": [{ "w:p": [{ "w:r": [{ "w:t": {} }] }] }] },
                        { "w:showingPlcHdr": {} },
                        { "w:text": { _attr: { "w:multiLine": false } } },
                    ],
                },
            ],
        });
    });
});

describe("StructuredDocumentTagProperties backward compatibility", () => {
    it("should accept a plain string alias", () => {
        const tree = new Formatter().format(new StructuredDocumentTagProperties("TOC"));
        expect(tree).to.deep.equal({
            "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "TOC" } } }],
        });
    });

    it("should accept undefined (empty properties)", () => {
        const tree = new Formatter().format(new StructuredDocumentTagProperties());
        expect(tree).to.deep.equal({ "w:sdtPr": {} });
    });
});

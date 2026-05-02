import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { MovedFromTextRun, MovedToTextRun } from "./moved-text-run";

describe("MovedFromTextRun", () => {
    it("should create a moveFrom run with text", () => {
        const tree = new Formatter().format(
            new MovedFromTextRun({
                id: 1,
                author: "John",
                date: "2024-01-01T00:00:00Z",
                text: "moved content",
            }),
        );
        expect(tree).to.deep.equal({
            "w:moveFrom": [
                {
                    _attr: {
                        "w:author": "John",
                        "w:date": "2024-01-01T00:00:00Z",
                        "w:id": 1,
                    },
                },
                {
                    "w:r": [
                        {
                            "w:t": [{ _attr: { "xml:space": "preserve" } }, "moved content"],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a moveFrom run with formatting", () => {
        const tree = new Formatter().format(
            new MovedFromTextRun({
                id: 2,
                author: "Jane",
                date: "2024-01-01T00:00:00Z",
                text: "formatted",
                bold: true,
            }),
        );
        expect(tree).to.deep.equal({
            "w:moveFrom": [
                {
                    _attr: {
                        "w:author": "Jane",
                        "w:date": "2024-01-01T00:00:00Z",
                        "w:id": 2,
                    },
                },
                {
                    "w:r": [
                        {
                            "w:rPr": [{ "w:b": {} }, { "w:bCs": {} }],
                        },
                        {
                            "w:t": [{ _attr: { "xml:space": "preserve" } }, "formatted"],
                        },
                    ],
                },
            ],
        });
    });
});

describe("MovedToTextRun", () => {
    it("should create a moveTo run with text", () => {
        const tree = new Formatter().format(
            new MovedToTextRun({
                id: 1,
                author: "John",
                date: "2024-01-01T00:00:00Z",
                text: "moved content",
            }),
        );
        expect(tree).to.deep.equal({
            "w:moveTo": [
                {
                    _attr: {
                        "w:author": "John",
                        "w:date": "2024-01-01T00:00:00Z",
                        "w:id": 1,
                    },
                },
                {
                    "w:r": [
                        {
                            "w:t": [{ _attr: { "xml:space": "preserve" } }, "moved content"],
                        },
                    ],
                },
            ],
        });
    });

    it("should create a moveTo run with formatting", () => {
        const tree = new Formatter().format(
            new MovedToTextRun({
                id: 2,
                author: "Jane",
                date: "2024-01-01T00:00:00Z",
                text: "formatted",
                italics: true,
            }),
        );
        expect(tree).to.deep.equal({
            "w:moveTo": [
                {
                    _attr: {
                        "w:author": "Jane",
                        "w:date": "2024-01-01T00:00:00Z",
                        "w:id": 2,
                    },
                },
                {
                    "w:r": [
                        {
                            "w:rPr": [{ "w:i": {} }, { "w:iCs": {} }],
                        },
                        {
                            "w:t": [{ _attr: { "xml:space": "preserve" } }, "formatted"],
                        },
                    ],
                },
            ],
        });
    });
});

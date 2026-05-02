import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import {
    MoveFromRangeEnd,
    MoveFromRangeStart,
    MoveToRangeEnd,
    MoveToRangeStart,
} from "./move-bookmark";

describe("MoveFromRangeStart", () => {
    it("should create moveFromRangeStart with id only", () => {
        const tree = new Formatter().format(new MoveFromRangeStart(1));
        expect(tree).to.deep.equal({
            "w:moveFromRangeStart": { _attr: { "w:id": 1 } },
        });
    });

    it("should create moveFromRangeStart with all attributes", () => {
        const tree = new Formatter().format(
            new MoveFromRangeStart(1, "moved1", "John", "2024-01-01T00:00:00Z"),
        );
        expect(tree).to.deep.equal({
            "w:moveFromRangeStart": {
                _attr: {
                    "w:id": 1,
                    "w:name": "moved1",
                    "w:author": "John",
                    "w:date": "2024-01-01T00:00:00Z",
                },
            },
        });
    });
});

describe("MoveFromRangeEnd", () => {
    it("should create moveFromRangeEnd", () => {
        const tree = new Formatter().format(new MoveFromRangeEnd(1));
        expect(tree).to.deep.equal({
            "w:moveFromRangeEnd": { _attr: { "w:id": 1 } },
        });
    });
});

describe("MoveToRangeStart", () => {
    it("should create moveToRangeStart with id only", () => {
        const tree = new Formatter().format(new MoveToRangeStart(2));
        expect(tree).to.deep.equal({
            "w:moveToRangeStart": { _attr: { "w:id": 2 } },
        });
    });

    it("should create moveToRangeStart with all attributes", () => {
        const tree = new Formatter().format(
            new MoveToRangeStart(2, "moved1", "John", "2024-01-01T00:00:00Z"),
        );
        expect(tree).to.deep.equal({
            "w:moveToRangeStart": {
                _attr: {
                    "w:id": 2,
                    "w:name": "moved1",
                    "w:author": "John",
                    "w:date": "2024-01-01T00:00:00Z",
                },
            },
        });
    });
});

describe("MoveToRangeEnd", () => {
    it("should create moveToRangeEnd", () => {
        const tree = new Formatter().format(new MoveToRangeEnd(2));
        expect(tree).to.deep.equal({
            "w:moveToRangeEnd": { _attr: { "w:id": 2 } },
        });
    });
});

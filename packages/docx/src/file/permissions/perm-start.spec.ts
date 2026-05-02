import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { EditGroupType, PermEnd, PermStart } from "./perm-start";

describe("PermStart", () => {
    it("should create a permStart with id only", () => {
        const tree = new Formatter().format(new PermStart({ id: "range1" }));
        expect(tree).to.deep.equal({
            "w:permStart": { _attr: { "w:id": "range1" } },
        });
    });

    it("should create a permStart with edGroup", () => {
        const tree = new Formatter().format(
            new PermStart({ id: "range1", edGroup: EditGroupType.EVERYONE }),
        );
        expect(tree).to.deep.equal({
            "w:permStart": {
                _attr: {
                    "w:id": "range1",
                    "w:edGrp": "everyone",
                },
            },
        });
    });

    it("should create a permStart with ed (individual user)", () => {
        const tree = new Formatter().format(
            new PermStart({ id: "range2", ed: "john@example.com" }),
        );
        expect(tree).to.deep.equal({
            "w:permStart": {
                _attr: {
                    "w:id": "range2",
                    "w:ed": "john@example.com",
                },
            },
        });
    });

    it("should create a permStart with column range", () => {
        const tree = new Formatter().format(
            new PermStart({ id: "range3", colFirst: 1, colLast: 3 }),
        );
        expect(tree).to.deep.equal({
            "w:permStart": {
                _attr: {
                    "w:id": "range3",
                    "w:colFirst": 1,
                    "w:colLast": 3,
                },
            },
        });
    });
});

describe("PermEnd", () => {
    it("should create a permEnd", () => {
        const tree = new Formatter().format(new PermEnd("range1"));
        expect(tree).to.deep.equal({
            "w:permEnd": { _attr: { "w:id": "range1" } },
        });
    });
});

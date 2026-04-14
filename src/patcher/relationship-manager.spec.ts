import { TargetModeType } from "@file/relationships/relationship/relationship";
import { describe, expect, it } from "vite-plus/test";

import { appendRelationship, getNextRelationshipIndex } from "./relationship-manager";

describe("relationship-manager", () => {
    describe("getNextRelationshipIndex", () => {
        it("should get next relationship index", () => {
            const output = getNextRelationshipIndex({
                elements: [
                    {
                        elements: [
                            { attributes: { Id: "rId1" }, name: "Relationship", type: "element" },
                            { attributes: { Id: "rId1" }, name: "Relationship", type: "element" },
                        ],
                        name: "Relationships",
                        type: "element",
                    },
                ],
            });
            expect(output).to.deep.equal(2);
        });

        it("should work with an empty relationship Id", () => {
            const output = getNextRelationshipIndex({
                elements: [
                    {
                        elements: [{ name: "Relationship", type: "element" }],
                        name: "Relationships",
                        type: "element",
                    },
                ],
            });
            expect(output).to.deep.equal(1);
        });

        it("should work with no relationships", () => {
            const output = getNextRelationshipIndex({
                elements: [
                    {
                        elements: [],
                        name: "Relationships",
                        type: "element",
                    },
                ],
            });
            expect(output).to.deep.equal(1);
        });
    });

    describe("appendRelationship", () => {
        it("should append a relationship", () => {
            const output = appendRelationship(
                {
                    elements: [
                        {
                            elements: [
                                {
                                    attributes: { Id: "rId1" },
                                    name: "Relationship",
                                    type: "element",
                                },
                                {
                                    attributes: { Id: "rId1" },
                                    name: "Relationship",
                                    type: "element",
                                },
                            ],
                            name: "Relationships",
                            type: "element",
                        },
                    ],
                },
                1,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "test",
                TargetModeType.EXTERNAL,
            );
            expect(output).to.deep.equal([
                { attributes: { Id: "rId1" }, name: "Relationship", type: "element" },
                { attributes: { Id: "rId1" }, name: "Relationship", type: "element" },
                {
                    attributes: {
                        Id: "rId1",
                        Target: "test",
                        TargetMode: TargetModeType.EXTERNAL,
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    },
                    name: "Relationship",
                    type: "element",
                },
            ]);
        });
    });
});

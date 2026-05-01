import { describe, expect, it } from "vite-plus/test";

import {
    createTextElementContents,
    getFirstLevelElements,
    patchSpaceAttribute,
    toJson,
} from "./util";

describe("util", () => {
    describe("toJson", () => {
        it("should return an Element", () => {
            const output = toJson("<xml></xml>");
            expect(output).to.be.an("object");
        });
    });

    describe("createTextElementContents", () => {
        it("should return an array of elements", () => {
            const output = createTextElementContents("hello");
            expect(output).to.deep.equal([{ text: "hello", type: "text" }]);
        });
    });

    describe("patchSpaceAttribute", () => {
        it("should return an element with the xml:space attribute", () => {
            const output = patchSpaceAttribute({ name: "xml", type: "element" });
            expect(output).to.deep.equal({
                attributes: {
                    "xml:space": "preserve",
                },
                name: "xml",
                type: "element",
            });
        });
    });

    describe("getFirstLevelElements", () => {
        it("should return an empty array if no elements are found", () => {
            const elements = getFirstLevelElements(
                { elements: [{ elements: [], name: "Relationships", type: "element" }] },
                "Relationships",
            );
            expect(elements).to.deep.equal([]);
        });

        it("should return an array if elements are found", () => {
            const elements = getFirstLevelElements(
                {
                    elements: [
                        {
                            elements: [{ name: "Relationship", type: "element" }],
                            name: "Relationships",
                            type: "element",
                        },
                    ],
                },
                "Relationships",
            );
            expect(elements).to.deep.equal([{ name: "Relationship", type: "element" }]);
        });
    });
});

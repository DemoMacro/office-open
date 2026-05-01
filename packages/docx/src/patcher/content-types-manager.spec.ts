import { describe, expect, it } from "vite-plus/test";

import { appendContentType } from "./content-types-manager";

describe("content-types-manager", () => {
    describe("appendContentType", () => {
        it("should append a content type", () => {
            const element = {
                elements: [
                    {
                        elements: [
                            {
                                type: "element",
                                name: "Default",
                            },
                        ],
                        name: "Types",
                        type: "element",
                    },
                ],
                name: "xml",
                type: "element",
            };
            appendContentType(
                element,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "docx",
            );

            expect(element).to.deep.equal({
                elements: [
                    {
                        elements: [
                            {
                                name: "Default",
                                type: "element",
                            },
                            {
                                attributes: {
                                    ContentType:
                                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                                    Extension: "docx",
                                },
                                name: "Default",
                                type: "element",
                            },
                        ],
                        name: "Types",
                        type: "element",
                    },
                ],
                name: "xml",
                type: "element",
            });
        });

        it("should not append duplicate content type", () => {
            const element = {
                elements: [
                    {
                        elements: [
                            {
                                type: "element",
                                name: "Default",
                                attributes: {
                                    ContentType: "image/png",
                                    Extension: "png",
                                },
                            },
                        ],
                        name: "Types",
                        type: "element",
                    },
                ],
                name: "xml",
                type: "element",
            };
            appendContentType(element, "image/png", "png");

            expect(element.elements.length).toBe(1);
        });
    });
});

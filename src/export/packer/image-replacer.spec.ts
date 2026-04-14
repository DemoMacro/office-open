import type { Media } from "@file/media";
import { describe, expect, it } from "vite-plus/test";

import { ImageReplacer } from "./image-replacer";

describe("ImageReplacer", () => {
    describe("#replace()", () => {
        it("should replace properly", () => {
            const imageReplacer = new ImageReplacer();
            const result = imageReplacer.replace(
                "test {test-image.png} test",
                [
                    {
                        data: Buffer.from(""),
                        fileName: "test-image.png",
                        transformation: {
                            emus: {
                                x: 100,
                                y: 100,
                            },
                            pixels: {
                                x: 100,
                                y: 100,
                            },
                        },
                        type: "png",
                    },
                ],
                0,
            );

            expect(result).to.equal("test 0 test");
        });
    });

    describe("#getMediaData()", () => {
        it("should get media data", () => {
            const imageReplacer = new ImageReplacer();
            const result = imageReplacer.getMediaData("test {test-image} test", {
                Array: [
                    {
                        dimensions: {
                            emus: {
                                x: 100,
                                y: 100,
                            },
                            pixels: {
                                x: 100,
                                y: 100,
                            },
                        },
                        fileName: "test-image",
                        stream: Buffer.from(""),
                    },
                ],
            } as unknown as Media);

            expect(result).to.have.length(1);
        });
    });
});

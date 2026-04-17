import * as convenienceFunctions from "@util/convenience-functions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { createTransformation, Media } from "./media";

describe("createTransformation", () => {
    it("should convert pixels to EMUs with default values", () => {
        const result = createTransformation({ width: 100, height: 200 });

        expect(result).to.deep.equal({
            emus: {
                x: 952500,
                y: 1905000,
            },
            flip: undefined,
            offset: {
                emus: {
                    x: 0,
                    y: 0,
                },
                pixels: {
                    x: 0,
                    y: 0,
                },
            },
            pixels: {
                x: 100,
                y: 200,
            },
            rotation: undefined,
        });
    });

    it("should convert offset from pixels to EMUs", () => {
        const result = createTransformation({
            height: 50,
            offset: { left: 10, top: 20 },
            width: 100,
        });

        expect(result.offset).to.deep.equal({
            emus: { x: 95250, y: 190500 },
            pixels: { x: 10, y: 20 },
        });
    });

    it("should convert rotation from degrees to 60000ths", () => {
        const result = createTransformation({
            height: 100,
            rotation: 90,
            width: 100,
        });

        expect(result.rotation).to.equal(5400000);
    });

    it("should pass flip through unchanged", () => {
        const result = createTransformation({
            flip: { horizontal: true, vertical: false },
            height: 100,
            width: 100,
        });

        expect(result.flip).to.deep.equal({ horizontal: true, vertical: false });
    });
});

describe("Media", () => {
    beforeEach(() => {
        vi.spyOn(convenienceFunctions, "uniqueId").mockReturnValue("test");
    });

    afterEach(() => {
        vi.resetAllMocks();
    });

    describe("#Array", () => {
        it("Get images as array", () => {
            const media = new Media();

            media.addImage("test2.png", {
                data: Buffer.from(""),
                fileName: "test.png",
                transformation: {
                    emus: {
                        x: Math.round(1 * 9525),
                        y: Math.round(1 * 9525),
                    },
                    flip: {
                        horizontal: true,
                        vertical: true,
                    },
                    pixels: {
                        x: Math.round(100),
                        y: Math.round(100),
                    },
                    rotation: 90,
                },
                type: "png",
            });

            const array = media.Array;
            expect(array).to.be.an.instanceof(Array);
            expect(array.length).to.equal(1);

            const image = array[0];
            expect(image.fileName).to.equal("test.png");
            expect(image.transformation).to.deep.equal({
                emus: {
                    x: 9525,
                    y: 9525,
                },
                flip: {
                    horizontal: true,
                    vertical: true,
                },
                pixels: {
                    x: 100,
                    y: 100,
                },
                rotation: 90,
            });
        });
    });
});

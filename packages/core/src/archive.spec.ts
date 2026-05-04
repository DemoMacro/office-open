import { describe, expect, it } from "vite-plus/test";

import { uint8ToBase64, getImageType } from "./archive";

describe("uint8ToBase64", () => {
    it("should convert empty array to empty string", () => {
        expect(uint8ToBase64(new Uint8Array(0))).toBe("");
    });

    it("should convert simple bytes to base64", () => {
        // "Hello" in bytes
        const bytes = new Uint8Array([72, 101, 108, 108, 111]);
        expect(uint8ToBase64(bytes)).toBe("SGVsbG8=");
    });
});

describe("getImageType", () => {
    it("should detect png", () => {
        expect(getImageType("photo.png")).toBe("png");
    });

    it("should detect jpg", () => {
        expect(getImageType("photo.jpg")).toBe("jpg");
    });

    it("should normalize jpeg to jpg", () => {
        expect(getImageType("photo.jpeg")).toBe("jpg");
    });

    it("should detect gif", () => {
        expect(getImageType("photo.gif")).toBe("gif");
    });

    it("should detect emf", () => {
        expect(getImageType("image.emf")).toBe("emf");
    });

    it("should default to png for unknown extension", () => {
        expect(getImageType("file.xyz")).toBe("png");
    });

    it("should handle path with directory", () => {
        expect(getImageType("word/media/image1.png")).toBe("png");
    });

    it("should handle uppercase extension", () => {
        expect(getImageType("photo.PNG")).toBe("png");
    });
});

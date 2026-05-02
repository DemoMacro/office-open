import { describe, expect, it } from "vite-plus/test";

import { hashedId, uniqueId, uniqueNumericIdCreator, uniqueUuid } from "./id-generators";

describe("uniqueNumericIdCreator", () => {
    it("should start from 1 by default", () => {
        const gen = uniqueNumericIdCreator();
        expect(gen()).toBe(1);
        expect(gen()).toBe(2);
        expect(gen()).toBe(3);
    });

    it("should start from a custom initial value", () => {
        const gen = uniqueNumericIdCreator(10);
        expect(gen()).toBe(11);
        expect(gen()).toBe(12);
    });

    it("should be independent across instances", () => {
        const gen1 = uniqueNumericIdCreator();
        const gen2 = uniqueNumericIdCreator(100);
        expect(gen1()).toBe(1);
        expect(gen2()).toBe(101);
        expect(gen1()).toBe(2);
    });
});

describe("uniqueId", () => {
    it("should return a lowercase string", () => {
        const id = uniqueId();
        expect(id).toBe(id.toLowerCase());
    });

    it("should return different values on each call", () => {
        const ids = new Set(Array.from({ length: 100 }, () => uniqueId()));
        expect(ids.size).toBe(100);
    });
});

describe("hashedId", () => {
    it("should return a consistent SHA-1 hex string", () => {
        const hash1 = hashedId("hello");
        const hash2 = hashedId("hello");
        expect(hash1).toBe(hash2);
        expect(hash1).toHaveLength(40); // SHA-1 = 20 bytes = 40 hex chars
    });

    it("should return different hashes for different inputs", () => {
        expect(hashedId("hello")).not.toBe(hashedId("world"));
    });

    it("should handle Buffer input", () => {
        const hash = hashedId(Buffer.from("hello"));
        expect(hash).toHaveLength(40);
    });

    it("should handle Uint8Array input", () => {
        const hash = hashedId(new TextEncoder().encode("hello"));
        expect(hash).toHaveLength(40);
    });
});

describe("uniqueUuid", () => {
    it("should return a UUID v4-style string (8-4-4-4-12 format)", () => {
        const uuid = uniqueUuid();
        expect(uuid).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/);
    });

    it("should return different values on each call", () => {
        const uuids = new Set(Array.from({ length: 100 }, () => uniqueUuid()));
        expect(uuids.size).toBe(100);
    });
});

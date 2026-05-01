import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createBlipFill } from "./blip-fill";

const createMockMediaData = (overrides: Record<string, unknown> = {}) =>
    ({
        fileName: "test.png",
        type: "png",
        transformation: { pixels: { x: 100, y: 100 }, emus: { x: 914400, y: 914400 } },
        ...overrides,
    }) as never;

describe("BlipFill", () => {
    describe("createBlipFill", () => {
        it("should default to stretch fill mode", () => {
            const blipFill = createBlipFill(createMockMediaData());
            const tree = new Formatter().format(blipFill);
            expect(tree).to.have.property("pic:blipFill");
            const children = tree["pic:blipFill"] as unknown[];
            expect(children.some((c) => c && "a:stretch" in (c as Record<string, unknown>))).toBe(
                true,
            );
        });

        it("should use tile fill mode when tile options provided", () => {
            const blipFill = createBlipFill(createMockMediaData(), { tile: {} });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            expect(children.some((c) => c && "a:tile" in (c as Record<string, unknown>))).toBe(
                true,
            );
        });

        it("should create tile with scale factors", () => {
            const blipFill = createBlipFill(createMockMediaData(), {
                tile: { sx: 50, sy: 50 },
            });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const tile = children.find((c) => c && "a:tile" in (c as Record<string, unknown>));
            expect(tile).toBeDefined();
            expect((tile as Record<string, unknown>)["a:tile"]).to.deep.equal({
                _attr: { sx: 50, sy: 50 },
            });
        });

        it("should create tile with flip mode", () => {
            const blipFill = createBlipFill(createMockMediaData(), {
                tile: { flip: "XY" },
            });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const tile = children.find((c) => c && "a:tile" in (c as Record<string, unknown>));
            expect((tile as Record<string, unknown>)["a:tile"]).to.deep.equal({
                _attr: { flip: "xy" },
            });
        });

        it("should create tile with alignment", () => {
            const blipFill = createBlipFill(createMockMediaData(), {
                tile: { align: "CENTER" },
            });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const tile = children.find((c) => c && "a:tile" in (c as Record<string, unknown>));
            expect((tile as Record<string, unknown>)["a:tile"]).to.deep.equal({
                _attr: { algn: "ctr" },
            });
        });

        it("should create tile with all options", () => {
            const blipFill = createBlipFill(createMockMediaData(), {
                tile: { tx: 100000, ty: 50000, sx: 50, sy: 50, flip: "X", align: "TOP_LEFT" },
            });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const tile = children.find((c) => c && "a:tile" in (c as Record<string, unknown>));
            expect((tile as Record<string, unknown>)["a:tile"]).to.deep.equal({
                _attr: {
                    algn: "tl",
                    flip: "x",
                    sx: 50,
                    sy: 50,
                    tx: 100000,
                    ty: 50000,
                },
            });
        });

        it("should include dpi attribute when specified", () => {
            const blipFill = createBlipFill(createMockMediaData(), { dpi: 300 });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const attr = children[0] as Record<string, unknown>;
            expect(attr).to.have.property("_attr");
            expect(attr._attr as Record<string, unknown>).to.have.property("dpi", 300);
        });

        it("should include rotWithShape attribute when specified", () => {
            const blipFill = createBlipFill(createMockMediaData(), { rotWithShape: true });
            const tree = new Formatter().format(blipFill);
            const children = tree["pic:blipFill"] as unknown[];
            const attr = children[0] as Record<string, unknown>;
            expect(attr).to.have.property("_attr");
            expect(attr._attr as Record<string, unknown>).to.have.property("rotWithShape", 1);
        });
    });
});

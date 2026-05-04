import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseTransform, parseGeometry, parseFill, parseOutline, parseEffects } from "./drawingml";

describe("parseTransform", () => {
    it("should parse position and size", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:xfrm",
                    elements: [
                        { name: "a:off", attributes: { x: "100", y: "200" } },
                        { name: "a:ext", attributes: { cx: "400", cy: "300" } },
                    ],
                },
            ],
        };
        const result = parseTransform(spPr);
        expect(result.x).toBe(100);
        expect(result.y).toBe(200);
        expect(result.width).toBe(400);
        expect(result.height).toBe(300);
    });

    it("should parse rotation (convert from 60000ths of degree)", () => {
        const spPr: Element = {
            elements: [{ name: "a:xfrm", attributes: { rot: "5400000" } }],
        };
        const result = parseTransform(spPr);
        expect(result.rotation).toBe(90);
    });

    it("should parse flip", () => {
        const spPr: Element = {
            elements: [{ name: "a:xfrm", attributes: { flipH: "1", flipV: "1" } }],
        };
        const result = parseTransform(spPr);
        expect(result.flipH).toBe(true);
        expect(result.flipV).toBe(true);
    });

    it("should return empty object for missing xfrm", () => {
        const result = parseTransform({ elements: [] });
        expect(result).toEqual({});
    });
});

describe("parseGeometry", () => {
    it("should parse preset geometry", () => {
        const spPr: Element = {
            elements: [{ name: "a:prstGeom", attributes: { prst: "ellipse" } }],
        };
        expect(parseGeometry(spPr)).toBe("ellipse");
    });

    it("should return undefined for missing geometry", () => {
        expect(parseGeometry({ elements: [] })).toBeUndefined();
    });
});

describe("parseFill", () => {
    it("should parse solid fill with srgbClr", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:solidFill",
                    elements: [{ name: "a:srgbClr", attributes: { val: "4472C4" } }],
                },
            ],
        };
        const result = parseFill(spPr);
        expect(result).toEqual({ type: "solid", color: "4472C4" });
    });

    it("should parse solid fill with schemeClr", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:solidFill",
                    elements: [{ name: "a:schemeClr", attributes: { val: "accent1" } }],
                },
            ],
        };
        const result = parseFill(spPr);
        expect(result).toEqual({ type: "solid", color: "accent1" });
    });

    it("should parse noFill", () => {
        const spPr: Element = {
            elements: [{ name: "a:noFill" }],
        };
        expect(parseFill(spPr)).toEqual({ type: "none" });
    });

    it("should return undefined for no fill elements", () => {
        expect(parseFill({ elements: [] })).toBeUndefined();
    });
});

describe("parseOutline", () => {
    it("should parse line properties", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:ln",
                    attributes: { w: "12700", cap: "rnd" },
                    elements: [
                        {
                            name: "a:solidFill",
                            elements: [{ name: "a:srgbClr", attributes: { val: "000000" } }],
                        },
                    ],
                },
            ],
        };
        const result = parseOutline(spPr);
        expect(result?.width).toBe(12700);
        expect(result?.cap).toBe("rnd");
        expect(result?.color).toBe("000000");
    });

    it("should parse dash style", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:ln",
                    elements: [{ name: "a:prstDash", attributes: { val: "dash" } }],
                },
            ],
        };
        const result = parseOutline(spPr);
        expect(result?.dashStyle).toBe("dash");
    });

    it("should parse noFill as width 0", () => {
        const spPr: Element = {
            elements: [{ name: "a:ln", elements: [{ name: "a:noFill" }] }],
        };
        const result = parseOutline(spPr);
        expect(result?.width).toBe(0);
    });

    it("should return undefined for missing ln", () => {
        expect(parseOutline({ elements: [] })).toBeUndefined();
    });
});

describe("parseEffects", () => {
    it("should parse outer shadow", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:effectLst",
                    elements: [
                        {
                            name: "a:outerShdw",
                            attributes: { blurRad: "50800", dist: "38100", dir: "5400000" },
                            elements: [
                                {
                                    name: "a:solidFill",
                                    elements: [
                                        { name: "a:srgbClr", attributes: { val: "000000" } },
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
        };
        const result = parseEffects(spPr);
        const shadow = result!.outerShadow as Record<string, unknown>;
        expect(shadow).toBeDefined();
        expect(shadow.blurRadius).toBe(50800);
        expect(shadow.distance).toBe(38100);
        expect(shadow.color).toBe("000000");
    });

    it("should parse glow", () => {
        const spPr: Element = {
            elements: [
                {
                    name: "a:effectLst",
                    elements: [{ name: "a:glow", attributes: { rad: "76200" } }],
                },
            ],
        };
        const result = parseEffects(spPr);
        const glow = result!.glow as Record<string, unknown>;
        expect(glow.radius).toBe(76200);
    });

    it("should return undefined for missing effectLst", () => {
        expect(parseEffects({ elements: [] })).toBeUndefined();
    });
});

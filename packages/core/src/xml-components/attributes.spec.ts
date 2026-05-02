import { describe, expect, it } from "vite-plus/test";

import { XmlAttributeComponent, NextAttributeComponent } from "./attributes";
import type { IContext } from "./base";

const emptyContext: IContext = { stack: [] };

describe("XmlAttributeComponent", () => {
    it("should map properties using xmlKeys", () => {
        class TestAttrs extends XmlAttributeComponent<{
            readonly val?: string;
            readonly color?: string;
        }> {
            protected readonly xmlKeys = {
                val: "w:val",
                color: "w:color",
            };
        }

        const attrs = new TestAttrs({ val: "Heading1", color: "FF0000" });
        expect(attrs.prepForXml(emptyContext)).toEqual({
            _attr: { "w:val": "Heading1", "w:color": "FF0000" },
        });
    });

    it("should skip undefined values", () => {
        class TestAttrs extends XmlAttributeComponent<{ readonly val?: string }> {
            protected readonly xmlKeys = { val: "w:val" };
        }

        const attrs = new TestAttrs({});
        expect(attrs.prepForXml(emptyContext)).toEqual({ _attr: {} });
    });

    it("should use property name as key when no xmlKeys mapping", () => {
        class TestAttrs extends XmlAttributeComponent<{ readonly customAttr: string }> {}

        const attrs = new TestAttrs({ customAttr: "value" });
        expect(attrs.prepForXml(emptyContext)).toEqual({
            _attr: { customAttr: "value" },
        });
    });
});

describe("NextAttributeComponent", () => {
    it("should create attributes with explicit key-value pairs", () => {
        const attr = new NextAttributeComponent({
            val: { key: "w:val", value: "Test" },
            id: { key: "w:id", value: 42 },
        });
        expect(attr.prepForXml(emptyContext)).toEqual({
            _attr: { "w:val": "Test", "w:id": 42 },
        });
    });

    it("should skip undefined values", () => {
        const attr = new NextAttributeComponent({
            val: { key: "w:val", value: undefined as unknown as string },
        });
        expect(attr.prepForXml(emptyContext)).toEqual({ _attr: {} });
    });
});

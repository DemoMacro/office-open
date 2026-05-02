import { describe, expect, it } from "vite-plus/test";

import type { IContext } from "./base";
import {
    BuilderElement,
    EmptyElement,
    HpsMeasureElement,
    NumberValueElement,
    OnOffElement,
    StringContainer,
    StringEnumValueElement,
    StringValueElement,
    chartAttr,
    wrapEl,
} from "./elements";

const emptyContext: IContext = { stack: [] };

describe("OnOffElement (CT_OnOff)", () => {
    it("should emit no val attribute when true (default)", () => {
        const el = new OnOffElement("w:b");
        expect(el.prepForXml(emptyContext)).toEqual({ "w:b": {} });
    });

    it("should emit no val attribute when explicitly true", () => {
        const el = new OnOffElement("w:b", true);
        expect(el.prepForXml(emptyContext)).toEqual({ "w:b": {} });
    });

    it("should emit w:val=false for w: element", () => {
        const el = new OnOffElement("w:b", false);
        expect(el.prepForXml(emptyContext)).toEqual({ "w:b": { _attr: { "w:val": false } } });
    });

    it("should emit m:val=false for m: element (dynamic namespace)", () => {
        const el = new OnOffElement("m:hideBot", false);
        expect(el.prepForXml(emptyContext)).toEqual({ "m:hideBot": { _attr: { "m:val": false } } });
    });

    it("should emit a:val=false for a: element", () => {
        const el = new OnOffElement("a:noAutofit", false);
        expect(el.prepForXml(emptyContext)).toEqual({
            "a:noAutofit": { _attr: { "a:val": false } },
        });
    });

    it("should handle element without namespace prefix", () => {
        const el = new OnOffElement("bold", false);
        expect(el.prepForXml(emptyContext)).toEqual({ bold: { _attr: { "bold:val": false } } });
    });
});

describe("HpsMeasureElement (CT_HpsMeasure)", () => {
    it("should accept a number", () => {
        const el = new HpsMeasureElement("w:sz", 24);
        const result = el.prepForXml(emptyContext);
        expect(result).toEqual({ "w:sz": { _attr: { "w:val": 24 } } });
    });

    it("should accept a universal measure string", () => {
        const el = new HpsMeasureElement("w:sz", "12pt");
        const result = el.prepForXml(emptyContext);
        expect(result).toEqual({ "w:sz": { _attr: { "w:val": "12pt" } } });
    });
});

describe("EmptyElement (CT_Empty)", () => {
    it("should produce an empty element", () => {
        const el = new EmptyElement("w:bookmarkStart");
        expect(el.prepForXml(emptyContext)).toEqual({ "w:bookmarkStart": {} });
    });
});

describe("StringValueElement (CT_String)", () => {
    it("should use w: namespace for w: element", () => {
        const el = new StringValueElement("w:pStyle", "Heading1");
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:pStyle": { _attr: { "w:val": "Heading1" } },
        });
    });

    it("should use m: namespace for m: element", () => {
        const el = new StringValueElement("m:mathFont", "Cambria Math");
        expect(el.prepForXml(emptyContext)).toEqual({
            "m:mathFont": { _attr: { "m:val": "Cambria Math" } },
        });
    });
});

describe("NumberValueElement", () => {
    it("should emit val attribute with number", () => {
        const el = new NumberValueElement("w:ilvl", 0);
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:ilvl": { _attr: { "w:val": 0 } },
        });
    });
});

describe("StringEnumValueElement", () => {
    it("should emit val attribute with enum string", () => {
        const el = new StringEnumValueElement("w:jc", "center");
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:jc": { _attr: { "w:val": "center" } },
        });
    });
});

describe("StringContainer", () => {
    it("should contain text content", () => {
        const el = new StringContainer("w:t", "Hello");
        expect(el.prepForXml(emptyContext)).toEqual({ "w:t": ["Hello"] });
    });
});

describe("BuilderElement", () => {
    it("should create an element with no attributes or children", () => {
        const el = new BuilderElement({ name: "w:p" });
        expect(el.prepForXml(emptyContext)).toEqual({ "w:p": {} });
    });

    it("should create an element with attributes (single attr unwrapped)", () => {
        const el = new BuilderElement({
            name: "w:pPr",
            attributes: {
                style: { key: "w:val", value: "Test" },
            },
        });
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:pPr": { _attr: { "w:val": "Test" } },
        });
    });

    it("should create an element with children", () => {
        const child = new EmptyElement("w:r");
        const el = new BuilderElement({ name: "w:p", children: [child] });
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:p": [{ "w:r": {} }],
        });
    });

    it("should create an element with both attributes and children", () => {
        const child = new StringContainer("w:t", "text");
        const el = new BuilderElement({
            name: "w:r",
            attributes: { lang: { key: "xml:lang", value: "en-US" } },
            children: [child],
        });
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:r": [{ _attr: { "xml:lang": "en-US" } }, { "w:t": ["text"] }],
        });
    });
});

describe("chartAttr", () => {
    it("should create attributes with explicit keys", () => {
        const attr = chartAttr({ "r:id": "rId1", "c:val": 42 });
        expect(attr.prepForXml(emptyContext)).toEqual({
            _attr: { "r:id": "rId1", "c:val": 42 },
        });
    });
});

describe("wrapEl", () => {
    it("should wrap a component in a named element", () => {
        const inner = new StringContainer("w:t", "text");
        const wrapped = wrapEl("w:r", inner);
        expect(wrapped.prepForXml(emptyContext)).toEqual({
            "w:r": [{ "w:t": ["text"] }],
        });
    });
});

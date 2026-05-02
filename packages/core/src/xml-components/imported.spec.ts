import { describe, expect, it } from "vite-plus/test";

import type { IContext } from "./base";
import { EMPTY_OBJECT } from "./component";
import {
    ImportedXmlComponent,
    ImportedRootElementAttributes,
    convertToXmlComponent,
} from "./imported";

const emptyContext: IContext = { stack: [] };

describe("ImportedXmlComponent", () => {
    it("should create from XML string", () => {
        const xml = '<w:p w:one="1"><w:r><w:t>Text</w:t></w:r></w:p>';
        const component = ImportedXmlComponent.fromXmlString(xml);
        expect(component).toBeDefined();

        const result = component.prepForXml(emptyContext);
        expect(result).toBeDefined();
        // rootKey comes from the XML root element
        const keys = Object.keys(result!);
        expect(keys).toContain("w:p");
    });

    it("should handle attributes", () => {
        const component = new ImportedXmlComponent("w:test", { attr1: "value1", attr2: "value2" });
        component.push(new ImportedXmlComponent("w:child"));

        const result = component.prepForXml(emptyContext);
        expect(result).toEqual({
            "w:test": [
                { _attr: { attr1: "value1", attr2: "value2" } },
                { "w:child": EMPTY_OBJECT },
            ],
        });
    });

    it("should handle text nodes", () => {
        const component = new ImportedXmlComponent("w:p");
        component.push("Hello");

        const result = component.prepForXml(emptyContext);
        expect(result).toEqual({ "w:p": ["Hello"] });
    });
});

describe("convertToXmlComponent", () => {
    it("should return undefined for invalid type", () => {
        expect(convertToXmlComponent({ type: "comment" } as any)).toBeUndefined();
    });

    it("should skip invalid children", () => {
        const element = {
            elements: [
                { text: "hello", type: "text" },
                { comment: "skip", type: "comment" } as any,
                { elements: [], name: "w:r", type: "element" },
            ],
            name: "w:p",
            type: "element",
        };
        const result = convertToXmlComponent(element);
        expect(result).toBeDefined();
        const json = JSON.parse(JSON.stringify(result));
        expect(json.root).toHaveLength(2);
    });
});

describe("ImportedRootElementAttributes", () => {
    it("should produce _attr output", () => {
        const attrs = new ImportedRootElementAttributes({ key: "value" });
        expect(attrs.prepForXml(emptyContext as IContext)).toEqual({ _attr: { key: "value" } });
    });
});

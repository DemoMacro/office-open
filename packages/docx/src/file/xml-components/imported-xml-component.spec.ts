import { EMPTY_OBJECT } from "@file/xml-components";
import {
    ImportedRootElementAttributes,
    ImportedXmlComponent,
    convertToXmlComponent,
} from "@office-open/core";
import { xml2js } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { beforeEach, describe, expect, it } from "vite-plus/test";

import type { IContext } from "./base";

const xmlString = `
        <w:p w:one="value 1" w:two="value 2">
            <w:rPr>
                <w:noProof>some value</w:noProof>
            </w:rPr>
            <w:r active="true">
                <w:t>Text 1</w:t>
            </w:r>
            <w:r active="true">
                <w:t>Text 2</w:t>
            </w:r>
        </w:p>
    `;

const convertedXmlElement = {
    root: [
        { root: { "w:one": "value 1", "w:two": "value 2" }, rootKey: "_attr" },
        { root: [{ rootKey: "w:noProof", root: ["some value"] }], rootKey: "w:rPr" },
        {
            root: [
                { rootKey: "_attr", root: { active: "true" } },
                { rootKey: "w:t", root: ["Text 1"] },
            ],
            rootKey: "w:r",
        },
        {
            root: [
                { rootKey: "_attr", root: { active: "true" } },
                { rootKey: "w:t", root: ["Text 2"] },
            ],
            rootKey: "w:r",
        },
    ],
    rootKey: "w:p",
};

describe("ImportedXmlComponent", () => {
    let importedXmlComponent: ImportedXmlComponent;

    beforeEach(() => {
        const attributes = {
            otherAttr: "2",
            someAttr: "1",
        };
        importedXmlComponent = new ImportedXmlComponent("w:test", attributes);

        importedXmlComponent.push(new ImportedXmlComponent("w:child"));
    });

    describe("#prepForXml()", () => {
        it("should transform for xml", () => {
            const converted = importedXmlComponent.prepForXml({ stack: [] } as unknown as IContext);
            expect(JSON.parse(JSON.stringify(converted))).to.deep.equal({
                "w:test": [
                    {
                        _attr: {
                            otherAttr: "2",
                            someAttr: "1",
                        },
                    },
                    {
                        "w:child": EMPTY_OBJECT,
                    },
                ],
            });
        });
    });

    it("should create XmlComponent from xml string", () => {
        const converted = ImportedXmlComponent.fromXmlString(xmlString);
        expect(JSON.parse(JSON.stringify(converted))).to.deep.equal(convertedXmlElement);
    });

    describe("convertToXmlComponent", () => {
        it("should convert to xml component", () => {
            const xmlObj = xml2js(xmlString, { compact: false }) as Element;
            const root = xmlObj.elements?.[0] ?? xmlObj;
            const converted = convertToXmlComponent(root as Element);
            expect(JSON.parse(JSON.stringify(converted))).to.deep.equal(convertedXmlElement);
        });

        it("should return undefined if xml type is invalid", () => {
            const xmlObj = { type: "invalid" } as Element;
            const converted = convertToXmlComponent(xmlObj);
            expect(converted).to.equal(undefined);
        });

        it("should skip child elements that return undefined", () => {
            const xmlObj: Element = {
                elements: [
                    { text: "hello", type: "text" },
                    { comment: "a comment", type: "comment" } as unknown as Element,
                    { elements: [], name: "w:r", type: "element" },
                ],
                name: "w:p",
                type: "element",
            };
            const converted = convertToXmlComponent(xmlObj);
            expect(converted).not.to.equal(undefined);
            // The comment child should be skipped (returns undefined)
            // We should have text and w:r but not the comment
            const json = JSON.parse(JSON.stringify(converted));
            expect(json.root).to.have.length(2);
        });
    });
});

describe("ImportedRootElementAttributes", () => {
    let attributes: ImportedRootElementAttributes;

    beforeEach(() => {
        attributes = new ImportedRootElementAttributes({});
    });

    describe("#prepForXml()", () => {
        it("should work", () => {
            const converted = attributes.prepForXml({} as IContext);
            expect(converted).to.deep.equal({
                _attr: {},
            });
        });
    });
});

import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentationSync } from "./generate";
import { parsePresentationSync } from "./parse";

const decoder = new TextDecoder();

describe("PPTX custom XML parts", () => {
  it("rebuilds items and property companions from typed options", () => {
    const output = generatePresentationSync(
      {
        slides: [],
        customXml: [{ content: "<item/>", properties: "<properties/>" }],
      },
      { type: "uint8array" },
    );
    const files = unzipSync(output);
    expect(decoder.decode(files["customXml/item1.xml"]!)).toContain("<item/>");
    expect(decoder.decode(files["customXml/itemProps1.xml"]!)).toContain("<properties/>");
    expect(decoder.decode(files["customXml/_rels/item1.xml.rels"]!)).toContain("customXmlProps");

    expect(parsePresentationSync(output).customXml).toEqual([
      {
        content: "<item/>",
        contentPath: "customXml/item1.xml",
        properties: "<properties/>",
        propertiesPath: "customXml/itemProps1.xml",
      },
    ]);
  });
});

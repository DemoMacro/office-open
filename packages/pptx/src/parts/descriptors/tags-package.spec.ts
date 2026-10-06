import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentationSync } from "../../generate";
import { parsePresentationSync } from "../../parse";

const decoder = new TextDecoder();

describe("user-defined tags part", () => {
  it("packs tags, wires customer data, and round-trips", () => {
    const options = {
      slides: [],
      tags: [{ name: "category", val: "demo" }],
    };
    const output = generatePresentationSync(options, { type: "uint8array" });
    const files = unzipSync(output);
    expect(decoder.decode(files["ppt/tags/tags1.xml"]!)).toContain(
      '<p:tag name="category" val="demo"/>',
    );
    expect(decoder.decode(files["ppt/presentation.xml"]!)).toContain("<p:tags");
    expect(decoder.decode(files["ppt/_rels/presentation.xml.rels"]!)).toContain(
      'Target="tags/tags1.xml"',
    );
    expect(decoder.decode(files["[Content_Types].xml"]!)).toContain("presentationml.tags+xml");

    expect(parsePresentationSync(output).tags).toEqual(options.tags);
  });

  it("round-trips an explicitly empty tag list", () => {
    const output = generatePresentationSync({ slides: [], tags: [] }, { type: "uint8array" });
    const files = unzipSync(output);
    expect(decoder.decode(files["ppt/tags/tags1.xml"]!)).toContain("<p:tagLst");
    expect(parsePresentationSync(output).tags).toEqual([]);
  });
});

import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { tagListDesc } from "./tags";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

describe("tagListDesc", () => {
  it("round-trips tags and escapes attribute values", () => {
    const tags = [
      { name: "category", val: "demo" },
      { name: "label", val: `A & B <C>` },
    ];
    const xml = tagListDesc.stringify(tags, writeCtx)!;
    expect(xml).toContain(`<p:tag name="category" val="demo"/>`);
    expect(xml).toContain(`val="A &amp; B &lt;C&gt;"`);
    expect(tagListDesc.parse(parseXml(xml).elements![0]!, readCtx)).toEqual(tags);
  });

  it("rejects malformed tags structurally", () => {
    const missing = parseXml('<p:tagLst><p:tag name="category"/></p:tagLst>');
    expect(() => tagListDesc.parse(missing.elements![0]!, readCtx)).toThrowError(
      /requires name and val/,
    );

    const unsupported = parseXml("<p:tagLst><p:other/></p:tagLst>");
    expect(() => tagListDesc.parse(unsupported.elements![0]!, readCtx)).toThrowError(
      /Unsupported tags child/,
    );
  });
});

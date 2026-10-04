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
});

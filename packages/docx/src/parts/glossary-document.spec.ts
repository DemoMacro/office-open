import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import { glossaryDesc } from "./glossary-document";
import type { GlossaryDocumentOptions } from "./glossary-document";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) => (typeof child === "string" ? child : "<w:p/>"),
  fileData: {} as never,
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: GlossaryDocumentOptions) {
  const xml = glossaryDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return glossaryDesc.parse(el, readCtx);
}

describe("glossaryDesc round-trip", () => {
  it("round-trips a simple building block", () => {
    const result = roundTrip({
      parts: [
        {
          name: "TestBlock",
          gallery: "default",
          children: [],
        },
      ],
    });
    expect(result.parts).toHaveLength(1);
    expect(result.parts[0].name).toBe("TestBlock");
    expect(result.parts[0].gallery).toBe("default");
  });

  it("round-trips category and gallery", () => {
    const result = roundTrip({
      parts: [
        {
          name: "CoverPage",
          gallery: "coverPg",
          category: "Built-In",
          children: [],
        },
      ],
    });
    expect(result.parts[0].gallery).toBe("coverPg");
    expect(result.parts[0].category).toBe("Built-In");
  });

  it("round-trips types", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Typed",
          gallery: "default",
          types: ["normal", "autoExp"],
          children: [],
        },
      ],
    });
    expect(result.parts[0].types).toEqual(["normal", "autoExp"]);
  });

  it("round-trips behaviors", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Behaved",
          gallery: "default",
          behaviors: ["content", "p"],
          children: [],
        },
      ],
    });
    expect(result.parts[0].behaviors).toEqual(["content", "p"]);
  });

  it("round-trips description", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Described",
          gallery: "default",
          description: "A test building block",
          children: [],
        },
      ],
    });
    expect(result.parts[0].description).toBe("A test building block");
  });

  it("round-trips guid", () => {
    const result = roundTrip({
      parts: [
        {
          name: "Guided",
          gallery: "default",
          guid: "12345678-ABCD-EF01-2345-6789ABCDEF01",
          children: [],
        },
      ],
    });
    expect(result.parts[0].guid).toBe("12345678-ABCD-EF01-2345-6789ABCDEF01");
  });

  it("round-trips multiple parts", () => {
    const result = roundTrip({
      parts: [
        { name: "Part1", gallery: "default", children: [] },
        { name: "Part2", gallery: "hdrs", children: [] },
      ],
    });
    expect(result.parts).toHaveLength(2);
    expect(result.parts[0].name).toBe("Part1");
    expect(result.parts[1].name).toBe("Part2");
  });

  it("round-trips empty parts", () => {
    const result = roundTrip({ parts: [] });
    expect(result.parts).toHaveLength(0);
  });
});

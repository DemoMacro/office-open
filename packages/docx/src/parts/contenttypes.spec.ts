import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ContentTypesInput } from "./contenttypes";
import { contentTypesDesc } from "./contenttypes";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: ContentTypesInput) {
  const xml = contentTypesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return contentTypesDesc.parse(el, readCtx);
}

describe("contentTypesDesc round-trip", () => {
  it("round-trips defaults", () => {
    const result = roundTrip({
      defaults: [
        { extension: "png", contentType: "image/png" },
        { extension: "xml", contentType: "application/xml" },
      ],
      overrides: [],
    });
    expect(result.defaults).toHaveLength(2);
    expect(result.defaults[0].extension).toBe("png");
    expect(result.defaults[0].contentType).toBe("image/png");
    expect(result.defaults[1].extension).toBe("xml");
  });

  it("round-trips overrides", () => {
    const result = roundTrip({
      defaults: [],
      overrides: [
        {
          partName: "/word/document.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        },
      ],
    });
    expect(result.overrides).toHaveLength(1);
    expect(result.overrides[0].partName).toBe("/word/document.xml");
    expect(result.overrides[0].contentType).toContain("wordprocessingml");
  });

  it("round-trips both defaults and overrides", () => {
    const result = roundTrip({
      defaults: [
        {
          extension: "rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ],
      overrides: [
        {
          partName: "/word/styles.xml",
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        },
        {
          partName: "/word/settings.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        },
      ],
    });
    expect(result.defaults).toHaveLength(1);
    expect(result.overrides).toHaveLength(2);
  });

  it("round-trips empty content types", () => {
    const result = roundTrip({ defaults: [], overrides: [] });
    expect(result.defaults).toHaveLength(0);
    expect(result.overrides).toHaveLength(0);
  });
});

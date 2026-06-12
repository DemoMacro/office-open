import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { relationshipsDesc } from "./relationships";
import type { RelationshipsInput } from "./relationships";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: RelationshipsInput) {
  const xml = relationshipsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return relationshipsDesc.parse(el, readCtx);
}

describe("relationshipsDesc round-trip", () => {
  it("round-trips a single relationship", () => {
    const result = roundTrip({
      relationships: [
        {
          id: 1,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
          target: "styles.xml",
        },
      ],
    });
    expect(result.relationships).toHaveLength(1);
    expect(result.relationships[0].id).toBe(1);
    expect(result.relationships[0].type).toBe(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    );
    expect(result.relationships[0].target).toBe("styles.xml");
  });

  it("round-trips multiple relationships", () => {
    const result = roundTrip({
      relationships: [
        { id: 1, type: "type-a", target: "a.xml" },
        { id: 2, type: "type-b", target: "b.xml" },
        { id: 3, type: "type-c", target: "c.xml" },
      ],
    });
    expect(result.relationships).toHaveLength(3);
    expect(result.relationships[0].target).toBe("a.xml");
    expect(result.relationships[1].target).toBe("b.xml");
    expect(result.relationships[2].target).toBe("c.xml");
  });

  it("round-trips relationship with targetMode", () => {
    const result = roundTrip({
      relationships: [
        {
          id: 5,
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
          target: "https://example.com",
          targetMode: "External",
        },
      ],
    });
    expect(result.relationships[0].targetMode).toBe("External");
    expect(result.relationships[0].target).toBe("https://example.com");
  });

  it("round-trips empty relationships", () => {
    const result = roundTrip({ relationships: [] });
    expect(result.relationships).toHaveLength(0);
  });

  it("round-trips string id", () => {
    const result = roundTrip({
      relationships: [{ id: "rId42", type: "type-x", target: "x.xml" }],
    });
    // parse converts "rId42" -> 42 (numeric extraction)
    expect(result.relationships[0].type).toBe("type-x");
    expect(result.relationships[0].target).toBe("x.xml");
  });
});

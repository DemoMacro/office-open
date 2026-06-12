import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { CorePropertiesInput } from "./core-properties";
import { corePropertiesDesc } from "./core-properties";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: CorePropertiesInput) {
  const xml = corePropertiesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return corePropertiesDesc.parse(el, readCtx) as Record<string, unknown>;
}

describe("corePropertiesDesc round-trip", () => {
  it("round-trips title", () => {
    const result = roundTrip({ title: "My Document" });
    expect(result.title).toBe("My Document");
  });

  it("round-trips creator", () => {
    const result = roundTrip({ creator: "John Doe" });
    expect(result.creator).toBe("John Doe");
  });

  it("round-trips subject", () => {
    const result = roundTrip({ subject: "Test Subject" });
    expect(result.subject).toBe("Test Subject");
  });

  it("round-trips keywords", () => {
    const result = roundTrip({ keywords: "test, docx, parse" });
    expect(result.keywords).toBe("test, docx, parse");
  });

  it("round-trips description", () => {
    const result = roundTrip({ description: "A test document" });
    expect(result.description).toBe("A test document");
  });

  it("round-trips lastModifiedBy", () => {
    const result = roundTrip({ lastModifiedBy: "Jane" });
    expect(result.lastModifiedBy).toBe("Jane");
  });

  it("round-trips revision", () => {
    const result = roundTrip({ revision: 3 });
    expect(result.revision).toBe(3);
  });

  it("round-trips all properties", () => {
    const result = roundTrip({
      title: "Title",
      subject: "Subject",
      creator: "Creator",
      keywords: "kw",
      description: "Desc",
      lastModifiedBy: "Modifier",
      revision: 5,
    });
    expect(result.title).toBe("Title");
    expect(result.subject).toBe("Subject");
    expect(result.creator).toBe("Creator");
    expect(result.keywords).toBe("kw");
    expect(result.description).toBe("Desc");
    expect(result.lastModifiedBy).toBe("Modifier");
    expect(result.revision).toBe(5);
  });

  it("round-trips empty properties", () => {
    const result = roundTrip({});
    expect(Object.keys(result).length).toBe(0);
  });

  it("round-trips special XML characters", () => {
    const result = roundTrip({ title: 'A <B> & "C"' });
    expect(result.title).toBe('A <B> & "C"');
  });
});

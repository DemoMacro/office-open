import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { customPropertiesDesc } from "./custom-properties";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(properties: { name: string; value: string }[]) {
  const xml = customPropertiesDesc.stringify({ properties }, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return customPropertiesDesc.parse(el, readCtx);
}

describe("customPropertiesDesc round-trip", () => {
  it("round-trips single property", () => {
    const result = roundTrip([{ name: "Category", value: "Finance" }]);
    expect(result.properties).toHaveLength(1);
    expect(result.properties[0].name).toBe("Category");
    expect(result.properties[0].value).toBe("Finance");
  });

  it("round-trips multiple properties", () => {
    const result = roundTrip([
      { name: "Department", value: "Engineering" },
      { name: "Status", value: "Draft" },
      { name: "Version", value: "2" },
    ]);
    expect(result.properties).toHaveLength(3);
    expect(result.properties[0].name).toBe("Department");
    expect(result.properties[1].name).toBe("Status");
    expect(result.properties[2].name).toBe("Version");
  });

  it("round-trips empty properties", () => {
    const result = roundTrip([]);
    expect(result.properties).toHaveLength(0);
  });

  it("round-trips special characters in value", () => {
    const result = roundTrip([{ name: "Notes", value: 'A <B> & "C"' }]);
    expect(result.properties[0].value).toBe('A <B> & "C"');
  });
});

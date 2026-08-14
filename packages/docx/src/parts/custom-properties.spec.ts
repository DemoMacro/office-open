import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { customPropertiesDesc } from "./custom-properties";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(properties: { name: string; value: string | number | boolean | Date }[]) {
  const xml = customPropertiesDesc.stringify({ properties }, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return customPropertiesDesc.parse(el, readCtx);
}

describe("customPropertiesDesc round-trip", () => {
  it("round-trips single property", () => {
    const result = roundTrip([{ name: "Category", value: "Finance" }]);
    expect(result.properties).toHaveLength(1);
    expect(result.properties[0]?.name).toBe("Category");
    expect(result.properties[0]?.value).toBe("Finance");
  });

  it("round-trips multiple properties", () => {
    const result = roundTrip([
      { name: "Department", value: "Engineering" },
      { name: "Status", value: "Draft" },
      { name: "Version", value: "2" },
    ]);
    expect(result.properties).toHaveLength(3);
    expect(result.properties[0]?.name).toBe("Department");
    expect(result.properties[1]?.name).toBe("Status");
    expect(result.properties[2]?.name).toBe("Version");
  });

  it("round-trips empty properties", () => {
    const result = roundTrip([]);
    expect(result.properties).toHaveLength(0);
  });

  it("round-trips special characters in value", () => {
    const result = roundTrip([{ name: "Notes", value: 'A <B> & "C"' }]);
    expect(result.properties[0]?.value).toBe('A <B> & "C"');
  });

  it("round-trips typed values with their vt:* element", () => {
    const when = new Date("2024-06-01T12:00:00.000Z");
    const result = roundTrip([
      { name: "Count", value: 42 },
      { name: "Ratio", value: 1.5 },
      { name: "Reviewed", value: true },
      { name: "Due", value: when },
    ]);
    expect(result.properties[0]?.value).toBe(42);
    expect(result.properties[1]?.value).toBe(1.5);
    expect(result.properties[2]?.value).toBe(true);
    const due = result.properties[3]?.value;
    expect(due).toBeInstanceOf(Date);
    expect((due as Date).toISOString()).toBe(when.toISOString());

    const xml = customPropertiesDesc.stringify(
      {
        properties: [
          { name: "Count", value: 42 },
          { name: "Ratio", value: 1.5 },
          { name: "Reviewed", value: true },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain("<vt:i4>42</vt:i4>");
    expect(xml).toContain("<vt:r8>1.5</vt:r8>");
    expect(xml).toContain("<vt:bool>true</vt:bool>");
  });
});

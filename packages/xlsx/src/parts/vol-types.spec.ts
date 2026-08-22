import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { VolTypeOptions } from "./vol-types";
import { buildVolTypesXml, parseVolTypesEl } from "./vol-types";

function roundTrip(volTypes: VolTypeOptions[]): VolTypeOptions[] {
  const xml = buildVolTypesXml(volTypes);
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return parseVolTypesEl(el);
}

describe("volTypes part", () => {
  it("emits the sml namespace on the part root", () => {
    const xml = buildVolTypesXml([{ type: "realTimeData" }]);
    expect(xml).toContain(
      '<volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    );
  });

  it("round-trips topics, string topics and references", () => {
    const volTypes: VolTypeOptions[] = [
      {
        type: "realTimeData",
        mains: [
          {
            first: "RTDServer.ProgID|topic!",
            topics: [
              {
                value: "42",
                valueType: "str",
                stringTopics: ["label"],
                refs: [{ reference: "Sheet1!A1", sheetIndex: 0 }],
              },
            ],
          },
        ],
      },
    ];
    expect(roundTrip(volTypes)).toEqual(volTypes);
  });

  it("normalizes the default valueType 'n' away on round-trip", () => {
    const result = roundTrip([
      { mains: [{ first: "x", topics: [{ value: "1", valueType: "n" }] }] },
    ]);
    expect(result[0]?.mains?.[0]?.topics?.[0]?.valueType).toBeUndefined();
  });

  it("returns empty for an empty collection", () => {
    expect(buildVolTypesXml([])).toBe("");
  });
});

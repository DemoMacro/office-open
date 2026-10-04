import { describe, expect, it } from "vite-plus/test";

import { generateOdt, OdtParseError, parseOdt } from "./index";

describe("ODT codec", () => {
  it("round-trips a document through canonical DocumentOptions", () => {
    const parsed = parseOdt(generateOdt({ sections: [{ children: [{ paragraph: "ODT" }] }] }));
    expect(parsed.sections[0]?.children).toEqual([{ paragraph: { text: "ODT" } }]);
  });

  it("wraps invalid packages in OdtParseError", () => {
    expect(() => parseOdt(new Uint8Array([1, 2, 3]))).toThrow(OdtParseError);
  });
});

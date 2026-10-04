import { describe, expect, it } from "vite-plus/test";

import { generateOds, OdsParseError, parseOds } from "./index";

describe("ODS codec", () => {
  it("round-trips a workbook through canonical WorkbookOptions", () => {
    const parsed = parseOds(
      generateOds({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    expect(parsed.worksheets?.[0]?.rows?.[0]?.cells).toEqual([{ reference: "A1", value: "ODS" }]);
  });

  it("wraps invalid packages in OdsParseError", () => {
    expect(() => parseOds(new Uint8Array([1, 2, 3]))).toThrow(OdsParseError);
  });
});

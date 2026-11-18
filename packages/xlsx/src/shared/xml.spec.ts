import { describe, expect, it } from "vite-plus/test";

import { decodeSpreadsheetXml, escapeSpreadsheetXml } from "./xml";

describe("SpreadsheetML XML strings", () => {
  it("round-trips XML-incompatible control characters", () => {
    const value = "before\u0000\u0001after";
    const encoded = escapeSpreadsheetXml(value);
    expect(encoded).toBe("before_x0000__x0001_after");
    expect(decodeSpreadsheetXml(encoded)).toBe(value);
  });

  it("does not reinterpret a literal SpreadsheetML escape", () => {
    const value = "_x0001_";
    expect(decodeSpreadsheetXml(escapeSpreadsheetXml(value))).toBe(value);
  });
});

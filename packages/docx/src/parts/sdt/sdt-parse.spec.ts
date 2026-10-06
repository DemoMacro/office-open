import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseSdtProperties } from "./sdt-parse";

describe("parseSdtProperties", () => {
  it("reads date language from w:lid/w:val", () => {
    const doc = parse(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:date w:fullDate="2001-01-01T00:00:00Z"><w:lid w:val="en-US"/></w:date></w:sdtPr>',
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    expect(parseSdtProperties(root).date?.languageId).toBe("en-US");
  });
});

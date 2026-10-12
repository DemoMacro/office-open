import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseChild } from "./bridge";

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

describe("slide child bridge", () => {
  it("parses content parts without rawXml absorption", () => {
    const doc = parseXml(
      '<p:contentPart xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId12"/>',
    );
    const element = doc.elements?.[0];
    if (!element) throw new Error("expected a content part element");
    expect(parseChild(element, readCtx)).toEqual({ contentPart: { rId: "rId12" } });
  });
});

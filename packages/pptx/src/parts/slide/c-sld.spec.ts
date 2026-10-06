import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseCustDataLst, stringifyCustDataLst } from "./c-sld";

describe("cSld customer data", () => {
  it("round-trips custData and tags references in schema order", () => {
    const items = [{ rId: "rId1" }, { rId: "rId2", kind: "tags" as const }];
    const xml = stringifyCustDataLst(items);
    expect(xml).toBe(
      '<p:custDataLst><p:custData r:id="rId1"/><p:tags r:id="rId2"/></p:custDataLst>',
    );
    const element = parseXml(xml).elements?.[0];
    expect(parseCustDataLst(element)).toEqual(items);
  });
});

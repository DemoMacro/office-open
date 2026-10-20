import { describe, expect, it } from "vite-plus/test";

import { collectRawXmlRelationshipIds } from "./raw-relationship-ids";

describe("collectRawXmlRelationshipIds", () => {
  it("collects relationship references from raw XML", () => {
    const ids = collectRawXmlRelationshipIds({
      rawXml: '<w:object><v:shape r:id="rId7"/></w:object>',
    });
    expect(ids).toEqual(new Set(["rId7"]));
  });

  it("treats binary views as leaves", () => {
    const binary = Object.assign(new Uint8Array(1), {
      rawXml: '<w:object r:id="rId7"/>',
    });
    expect(collectRawXmlRelationshipIds(binary)).toEqual(new Set());
  });
});

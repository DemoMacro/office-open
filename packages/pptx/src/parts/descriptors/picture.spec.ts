import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { pictureDesc } from "./shape";

const readContext: ReadContext = {
  resolveRelationship: (relationshipId) => (relationshipId === "rIdBroken" ? "NULL" : undefined),
  getPart: () => undefined,
  getRaw: () => undefined,
};

class WriteContextStub {
  registerShapeId() {}
  addRelationship() {
    return "rId1";
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
}

describe("broken picture references", () => {
  it("preserves an explicit preferRelativeResize=false switch", () => {
    const options = {
      type: "png",
      relationshipId: "rIdBroken",
      preferRelativeResize: false,
    } as const;
    const emitted = pictureDesc.stringify(options, new WriteContextStub() as WriteContext)!;
    expect(emitted).toContain('<p:cNvPicPr preferRelativeResize="0"/>');

    const element = parseXml(emitted).elements?.[0];
    if (!element) throw new Error("fixture has no root element");
    expect(pictureDesc.parse(element, readContext).preferRelativeResize).toBe(false);
  });

  it("round-trips a:blip r:embed without media bytes", () => {
    const xml =
      '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:nvPicPr><p:cNvPr id="1" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdBroken"/><a:stretch><a:fillRect/></a:stretch></p:blipFill></p:pic>';
    const element = parseXml(xml).elements?.[0];
    if (!element) throw new Error("fixture has no root element");
    const parsed = pictureDesc.parse(element, readContext);
    expect(parsed.relationshipId).toBe("rIdBroken");
    expect(parsed.data).toBeUndefined();

    const emitted = pictureDesc.stringify(
      { type: "png", relationshipId: parsed.relationshipId },
      new WriteContextStub() as WriteContext,
    )!;
    expect(emitted).toContain('<a:blip r:embed="rIdBroken"/>');
  });
});

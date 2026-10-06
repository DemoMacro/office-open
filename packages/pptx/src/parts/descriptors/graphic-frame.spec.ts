import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { readNvPrPlaceholder, stringifyNvPr } from "./graphic-frame";

describe("graphic frame nvPr customer data", () => {
  it("round-trips p:tags in p:nvPr", () => {
    const options = { customerData: [{ rId: "rId3", kind: "tags" as const }] };
    const xml = stringifyNvPr(options);
    expect(xml).toContain('<p:custDataLst><p:tags r:id="rId3"/></p:custDataLst>');

    const element = parseXml(`<p:nvGraphicFramePr>${xml}</p:nvGraphicFramePr>`).elements?.[0];
    if (!element) throw new Error("fixture has no root element");
    const result: Record<string, unknown> = {};
    readNvPrPlaceholder(element, result);
    expect(result.customerData).toEqual(options.customerData);
  });
});

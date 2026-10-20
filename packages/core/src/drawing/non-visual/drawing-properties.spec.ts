import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import {
  parseNonVisualDrawingProperties,
  stringifyNonVisualDrawingProperties,
} from "./drawing-properties";

describe("non-visual drawing properties", () => {
  it("preserves an empty extension list", () => {
    const xml = '<cNvPr id="1" name="Shape"><a:extLst/></cNvPr>';
    const options = parseNonVisualDrawingProperties(parseXml(xml).elements?.[0]);
    const output = stringifyNonVisualDrawingProperties("cNvPr", 1, options, "Shape");
    expect(options.ext).toBe("");
    expect(output).toContain("<a:extLst");
  });

  it("round-trips creationId alongside another typed extension", () => {
    const xml =
      '<cNvPr id="1" name="Shape"><a:extLst>' +
      '<a:ext uri="{creation-uri}"><a16:creationId id="{id}"/></a:ext>' +
      '<a:ext uri="{compat-uri}"><a14:compatExt spid="_x0000_s1025"/></a:ext>' +
      "</a:extLst></cNvPr>";
    const element = parseXml(xml).elements?.[0];
    const options = parseNonVisualDrawingProperties(element);
    const output = stringifyNonVisualDrawingProperties("cNvPr", 1, options, "Shape");
    expect(output).toContain("{creation-uri}");
    expect(output).toContain("{compat-uri}");
  });
});

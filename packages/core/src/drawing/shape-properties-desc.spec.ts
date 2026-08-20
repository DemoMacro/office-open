import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ReadContext, WriteContext } from "../descriptor";
import { shapePropertiesDesc } from "./shape-properties-desc";

const HIDDEN_LINE_URI = "{91240B29-F687-4F45-9708-019B960494DF}";

function roundTrip(xml: string) {
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  const parsed = shapePropertiesDesc.parse(el, {} as ReadContext);
  const generated = shapePropertiesDesc.stringify(parsed, {} as WriteContext);
  return { parsed, generated };
}

describe("shapePropertiesDesc", () => {
  it("round-trips the Office 2010 hidden-line extension structurally", () => {
    const { parsed, generated } = roundTrip(
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
        `<a:extLst><a:ext uri="${HIDDEN_LINE_URI}">` +
        '<a14:hiddenLine w="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
        "<a:round/><a:headEnd/><a:tailEnd/></a14:hiddenLine>" +
        "</a:ext></a:extLst></a:spPr>",
    );

    expect(parsed.extensions).toEqual([
      {
        uri: HIDDEN_LINE_URI,
        hiddenLine: {
          width: 0,
          type: "solidFill",
          color: { value: "000000" },
          join: "round",
          headEnd: {},
          tailEnd: {},
        },
      },
    ]);
    expect(generated).toContain(`<a:ext uri="${HIDDEN_LINE_URI}">`);
    expect(generated).toContain(
      '<a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" w="0">',
    );
    expect(generated).toContain('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>');
    expect(generated).toContain("<a:round/>");
    expect(generated).toContain("<a:headEnd/>");
    expect(generated).toContain("<a:tailEnd/>");
  });

  it("retains unknown extension payloads alongside a hidden line", () => {
    const { parsed, generated } = roundTrip(
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" ' +
        'xmlns:x="urn:example">' +
        '<a:extLst><a:ext uri="urn:example"><x:payload val="1"/></a:ext>' +
        `<a:ext uri="${HIDDEN_LINE_URI}"><a14:hiddenLine/></a:ext></a:extLst>` +
        "</a:spPr>",
    );

    expect(parsed.extensions?.[0]).toEqual({
      uri: "urn:example",
      content: '<x:payload val="1"/>',
    });
    expect(generated).toContain('<a:ext uri="urn:example"><x:payload val="1"/></a:ext>');
    expect(generated).toContain(
      `<a:ext uri="${HIDDEN_LINE_URI}"><a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"/></a:ext>`,
    );
  });

  it("preserves an explicitly empty extension list", () => {
    const { parsed, generated } = roundTrip(
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:extLst/></a:spPr>',
    );

    expect(parsed.extensions).toBeUndefined();
    expect(parsed.ext).toBe("");
    expect(generated).toBe("<a:extLst></a:extLst>");
  });
});

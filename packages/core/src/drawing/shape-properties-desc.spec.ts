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
  if (!generated) throw new Error("shape properties did not stringify");
  return { parsed, generated };
}

describe("shapePropertiesDesc", () => {
  it("round-trips an explicit empty transform", () => {
    const source =
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:xfrm/></a:spPr>';
    const el = parseXml(source).elements?.[0];
    if (!el) throw new Error("expected shape properties root");
    const parsed = shapePropertiesDesc.parse(el, {} as ReadContext);
    expect(parsed.transformPresent).toBe(true);

    const xml = shapePropertiesDesc.stringify(parsed, {} as WriteContext);
    expect(xml).toContain("<a:xfrm/>");
  });
  it("round-trips a14:hiddenEffects", () => {
    const xml =
      '<a:spPr><a:extLst><a:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}">' +
      '<a14:hiddenEffects xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
      "<a:effectLst><a:outerShdw dist='35921'/></a:effectLst>" +
      "</a14:hiddenEffects></a:ext></a:extLst></a:spPr>";
    const { generated, parsed } = roundTrip(xml);
    expect(parsed.extensions?.[0]?.hiddenEffects?.outerShadow?.distance).toBe(35921);
    expect(generated).toContain(
      '<a14:hiddenEffects xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
        "<a:effectLst><a:outerShdw",
    );
  });

  it("round-trips the Office 2010 hidden-fill extension structurally", () => {
    const { parsed, generated } = roundTrip(
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
        '<a:extLst><a:ext uri="urn:example:hidden-fill">' +
        '<a14:hiddenFill><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a14:hiddenFill>' +
        "</a:ext></a:extLst></a:spPr>",
    );

    expect(parsed.extensions).toEqual([
      {
        uri: "urn:example:hidden-fill",
        hiddenFill: { type: "solid", color: { value: "000000" } },
      },
    ]);
    expect(generated).toContain(
      '<a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:solidFill>',
    );
    expect(generated).toContain("</a:solidFill></a14:hiddenFill>");
  });

  it("keeps mixed hidden extensions in separate a:ext children", () => {
    const { generated, parsed } = roundTrip(
      '<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
        '<a:extLst><a:ext uri="urn:example:hidden-fill">' +
        "<a14:hiddenFill><a:noFill/></a14:hiddenFill></a:ext>" +
        '<a:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}">' +
        "<a14:hiddenEffects><a:effectLst><a:outerShdw dist='35921'/></a:effectLst>" +
        "</a14:hiddenEffects></a:ext></a:extLst></a:spPr>",
    );
    expect(parsed.extensions).toHaveLength(2);
    expect(generated.match(/<a:ext /g)).toHaveLength(2);
    expect(generated).not.toContain("</a14:hiddenFill><a:ext");
    expect(generated).toContain("<a14:hiddenFill");
    expect(generated).toContain("<a:noFill/></a14:hiddenFill>");
  });

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

    expect(parsed.extensions).toEqual([]);
    expect(parsed.ext).toBeUndefined();
    expect(generated).toBe("<a:extLst></a:extLst>");
  });
});

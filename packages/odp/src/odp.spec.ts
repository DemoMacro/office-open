import { generateOcf, ODF_NAMESPACES } from "@office-open/ocf";
import { describe, expect, it } from "vite-plus/test";

import { generateOdp, OdpParseError, parseOdp, type OdpPresentationOptions } from "./index";

describe("ODP codec", () => {
  it("round-trips a presentation through canonical PresentationOptions", () => {
    const parsed = parseOdp(generateOdp({ slides: [{ notes: "ODP" }] }));
    expect(parsed.slides?.[0]?.notes).toBe("ODP");
  });

  it("wraps invalid packages in OdpParseError", () => {
    expect(() => parseOdp(new Uint8Array([1, 2, 3]))).toThrow(OdpParseError);
  });

  it("round-trips page-level forms", () => {
    const options: OdpPresentationOptions = {
      slides: [
        {
          forms: {
            forms: [{ name: "Slide form", controls: [{ kind: "form:button", id: "next" }] }],
          },
        },
      ],
    };
    expect(parseOdp(generateOdp(options)).slides?.[0]?.forms).toMatchObject(
      options.slides![0]!.forms ?? {},
    );
  });

  it("round-trips slide charts with position, size, semantics, and order", () => {
    const parsed = parseOdp(
      generateOdp({
        slides: [
          {
            children: [
              {
                chart: {
                  name: "First",
                  type: "column",
                  x: 360000,
                  y: 720000,
                  width: 2160000,
                  height: 1440000,
                  series: [{ name: "Sales", values: [1, 2] }],
                },
              },
              {
                chart: {
                  name: "Second",
                  type: "bar",
                  x: 1080000,
                  y: 1440000,
                  width: 1800000,
                  height: 1080000,
                  series: [{ name: "Margin", values: [2, 4] }],
                },
              },
            ],
          },
        ],
      }),
    );
    const children = parsed.slides?.[0]?.children ?? [];
    expect(children.map((child) => ("chart" in child ? child.chart.name : undefined))).toEqual([
      "First",
      "Second",
    ]);
    expect(children[0]).toMatchObject({
      chart: {
        name: "First",
        type: "column",
        x: 360000,
        y: 720000,
        width: 2160000,
        height: 1440000,
        series: [{ name: "Sales", values: [1, 2] }],
      },
    });
    expect(children[1]).toMatchObject({ chart: { type: "bar" } });
  });

  it("reports unknown slide elements with structured errors", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:unknown/></draw:page></office:presentation></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseOdp(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdpParseError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: "/office:document-content/office:body/office:presentation/draw:page/draw:unknown",
      name: "draw:unknown",
      reason: "element has no canonical PresentationOptions mapping",
    });
  });
});

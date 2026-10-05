import { generateOcf, ODF_NAMESPACES, readOcf } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import {
  generatePresentation,
  OdpParseError,
  parsePresentation,
  type OdpDocumentOptions,
} from "./index";

describe("ODP codec", () => {
  it("round-trips a presentation through canonical PresentationOptions", () => {
    const parsed = parsePresentation(generatePresentation({ slides: [{ notes: "ODP" }] }));
    expect(parsed.slides?.[0]?.notes).toBe("ODP");
  });

  it("round-trips a linked picture with a parent-relative URL", () => {
    const options: OdpDocumentOptions = {
      slides: [
        {
          children: [
            {
              picture: {
                type: "png",
                sourceUrl: "../images/logo.png",
                x: 360000,
                y: 360000,
                width: 360000,
                height: 360000,
              },
            },
          ],
        },
      ],
    };
    expect(
      parsePresentation(generatePresentation(options)).slides?.[0]?.children?.[0],
    ).toMatchObject({ picture: { sourceUrl: "../images/logo.png" } });
  });

  it("restores percent and length semantics in style overlays", () => {
    const options: OdpDocumentOptions = {
      slides: [{ notes: "Overlay" }],
      styleOverlays: [
        {
          name: "P1",
          family: "paragraph",
          properties: [
            {
              name: "style:paragraph-properties",
              attributes: { "fo:line-height": 150, "fo:margin-left": 360000 },
            },
          ],
        },
      ],
    };
    const generated = generatePresentation(options);
    const first = readOcf(generated, "application/vnd.oasis.opendocument.presentation").files[
      "content.xml"
    ]!;
    expect(first).toContain('fo:line-height="150%"');
    expect(first).toContain('fo:margin-left="1cm"');

    const parsed = parsePresentation(generated) as OdpDocumentOptions;
    const second = readOcf(
      generatePresentation(parsed),
      "application/vnd.oasis.opendocument.presentation",
    ).files["content.xml"]!;
    expect(second).toContain('fo:line-height="150%"');
    expect(second).toContain('fo:margin-left="1cm"');
  });

  it("wraps invalid packages in OdpParseError", () => {
    expect(() => parsePresentation(new Uint8Array([1, 2, 3]))).toThrow(OdpParseError);
  });

  it("round-trips slide form containers and controls", () => {
    const forms = [
      {
        name: "Responses",
        automaticFocus: true,
        designMode: true,
        controls: [
          {
            control: "text" as const,
            id: "name",
            name: "Name",
            value: "Value",
            valueType: "string" as const,
            maxLength: 24,
            disabled: true,
            tabIndex: 2,
            automaticFocus: true,
          },
          { control: "checkBox" as const, id: "accepted", checked: true },
          {
            control: "dropDownList" as const,
            id: "choice",
            entries: ["One", "Two"],
            selectedIndex: 1,
          },
        ],
      },
    ];
    const parsed = parsePresentation(generatePresentation({ slides: [{ forms }] }));
    expect(parsed.slides?.[0]?.forms).toEqual(forms);
  });

  it("rejects unknown form controls with structured errors", () => {
    const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"><office:body><office:presentation><draw:page><office:forms><form:form><form:unknown/></form:form></office:forms></draw:page></office:presentation></office:body></office:document-content>`;
    let error: unknown;
    try {
      parsePresentation(
        generateOcf("application/vnd.oasis.opendocument.presentation", {
          "content.xml": content,
        }),
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdpParseError);
    expect(error).toMatchObject({ part: "content.xml", name: "form:unknown" });
  });

  it("round-trips slide charts with position, size, semantics, and order", () => {
    const parsed = parsePresentation(
      generatePresentation({
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
      parsePresentation(
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

import type { ChartSpaceOptions, ChartStyleResources } from "@office-open/core";
import { chartSpaceDesc, stringify, type WriteContext } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { pushChartStyle } from "./chart-style";
import { generateChartDocument, parseChartDocument } from "./index";

function style(fill: Record<string, string>, outline: Record<string, string> = {}) {
  return {
    name: "chart-style",
    attributes: { "style:family": "chart" },
    properties: [
      { name: "style:graphic-properties", attributes: { ...fill, ...outline } },
      { name: "style:chart-properties", attributes: {} },
    ],
  };
}

function chart(resources: ChartStyleResources): ChartSpaceOptions {
  return {
    type: "column",
    series: [{ name: "Sales", values: [1, 2] }],
    shapeProperties: { styleName: "chart-style" },
    styleResources: resources,
  };
}

async function parse(resources: ChartStyleResources) {
  return (await parseChartDocument(generateChartDocument({ chart: chart(resources) }))).chart;
}

describe("ODF chart style resources", () => {
  it("deduplicates generated fill resources and style names", () => {
    const resources: ChartStyleResources = {
      styles: [],
      gradients: [],
      hatches: [],
      fillImages: [],
      strokeDashes: [],
    };
    const fill = { type: "gradient", stops: [{ position: 0, color: "000000" }] } as const;
    const first = pushChartStyle({ fill }, resources, [], "series");
    const second = pushChartStyle({ fill }, resources, [], "other");
    expect(first).toBe(second);
    expect(resources.gradients).toHaveLength(1);
    expect(resources.styles).toHaveLength(0);
    expect(resources.hatches).toHaveLength(0);
    expect(resources.fillImages).toHaveLength(0);
  });

  it("projects all supported gradient kinds and scales intensity in RGB", async () => {
    const cases = [
      ["linear", undefined],
      ["axial", undefined],
      ["radial", "circle"],
      ["ellipsoid", "shape"],
      ["rectangular", "rect"],
    ] as const;
    for (const [kind, path] of cases) {
      const parsed = await parse({
        styles: [style({ "draw:fill": "gradient", "draw:fill-gradient-name": "gradient" })],
        gradients: [
          {
            name: "gradient",
            displayName: "Gradient",
            attributes: {
              "draw:style": kind,
              "draw:angle": "30",
              "draw:start-color": "#3366ff",
              "draw:start-intensity": "50%",
              "draw:end-color": "#ff0000",
              "draw:end-intensity": "25%",
            },
          },
        ],
        hatches: [],
        fillImages: [],
        strokeDashes: [],
      });
      expect(parsed.shapeProperties?.fill).toEqual({
        type: "gradient",
        angle: kind === "linear" || kind === "axial" ? 30 : undefined,
        ...(path ? { path } : {}),
        stops:
          kind === "axial"
            ? [
                { position: 0, color: "1A3380" },
                { position: 50, color: "400000" },
                { position: 100, color: "1A3380" },
              ]
            : [
                { position: 0, color: "1A3380" },
                { position: 100, color: "400000" },
              ],
      });
    }
  });

  it("preserves hatch styling and projects it for OOXML generation", async () => {
    const resources: ChartStyleResources = {
      styles: [style({ "draw:fill": "hatch", "draw:fill-hatch-name": "hatch" })],
      gradients: [],
      hatches: [
        {
          name: "hatch",
          displayName: "Hatch",
          attributes: {
            "draw:style": "single",
            "draw:color": "#00ff00",
            "draw:distance": "0.1cm",
            "draw:rotation": "45",
          },
        },
      ],
      fillImages: [],
      strokeDashes: [],
    };
    const parsed = await parse(resources);
    expect(parsed.shapeProperties?.styleName).toBe("chart-style");
    expect(stringify(chartSpaceDesc, parsed, {} as WriteContext)).toContain("<c:chartSpace");
  });

  it("projects embedded and linked bitmap fills", async () => {
    const image = {
      name: "image",
      displayName: "Image",
      attributes: { "xlink:href": "media/image.png", "svg:width": "2cm", "svg:height": "2cm" },
    };
    const embedded = await parse({
      styles: [
        style({
          "draw:fill": "bitmap",
          "draw:fill-image-name": "image",
          "style:repeat": "repeat",
          "draw:fill-image-width": "1cm",
          "draw:fill-image-height": "1cm",
          "draw:fill-image-ref-point": "top-left",
        }),
      ],
      gradients: [],
      hatches: [],
      fillImages: [{ ...image, data: new Uint8Array([1, 2, 3]) }],
      strokeDashes: [],
    });
    expect(embedded.shapeProperties?.fill).toMatchObject({
      type: "blip",
      imageType: "png",
      stretch: false,
      tile: { sx: 50, sy: 50, alignment: "topLeft" },
    });

    const linked = await parse({
      styles: [
        style({
          "draw:fill": "bitmap",
          "draw:fill-image-name": "image",
          "style:repeat": "stretch",
        }),
      ],
      gradients: [],
      hatches: [],
      fillImages: [
        { name: "image", attributes: { "xlink:href": "https://example.invalid/image.png" } },
      ],
      strokeDashes: [],
    });
    expect(linked.shapeProperties?.fill).toMatchObject({
      type: "blip",
      linkedUrl: "https://example.invalid/image.png",
    });
  });

  it("rejects unsupported no-repeat bitmap projection", async () => {
    await expect(
      parse({
        styles: [
          style({
            "draw:fill": "bitmap",
            "draw:fill-image-name": "image",
            "style:repeat": "no-repeat",
            "draw:tile-repeat-offset": "10% horizontal",
            "draw:opacity-name": "opaque",
          }),
        ],
        gradients: [],
        hatches: [],
        fillImages: [
          { name: "image", attributes: { "xlink:href": "https://example.invalid/image.png" } },
        ],
        strokeDashes: [],
      }),
    ).rejects.toMatchObject({ reason: "unsupported-bitmap-repeat" });
  });

  it("validates repeat tokens for source-reference bitmap fills", async () => {
    await expect(
      parse({
        styles: [
          style({
            "draw:fill": "bitmap",
            "draw:fill-image-name": "image",
            "style:repeat": "sideways",
          }),
        ],
        gradients: [],
        hatches: [],
        fillImages: [{ name: "image", attributes: { "xlink:href": "#source" } }],
        strokeDashes: [],
      }),
    ).rejects.toMatchObject({ reason: "unknown-bitmap-repeat" });
  });

  it("requires known bitmap repeat tokens", async () => {
    const resources = {
      styles: [
        style({
          "draw:fill": "bitmap",
          "draw:fill-image-name": "image",
          "style:repeat": "sideways",
        }),
      ],
      gradients: [],
      hatches: [],
      fillImages: [
        {
          name: "image",
          attributes: { "xlink:href": "media/image.png" },
          data: new Uint8Array([1]),
        },
      ],
      strokeDashes: [],
    } as ChartStyleResources;
    await expect(parse(resources)).rejects.toMatchObject({ reason: "unknown-bitmap-repeat" });
  });

  it("preserves square gradients and round dash caps without projecting them", async () => {
    const square = await parse({
      styles: [style({ "draw:fill": "gradient", "draw:fill-gradient-name": "gradient" })],
      gradients: [{ name: "gradient", attributes: { "draw:style": "square" } }],
      hatches: [],
      fillImages: [],
      strokeDashes: [],
    });
    expect(square.shapeProperties?.fill).toEqual({ type: "gradient", stops: [] });
    expect(square.shapeProperties?.styleName).toBe("chart-style");

    const roundDash = await parse({
      styles: [
        style(
          { "draw:fill": "none" },
          {
            "draw:stroke": "dash",
            "draw:stroke-dash": "dash",
            "svg:stroke-width": "0.1cm",
          },
        ),
      ],
      gradients: [],
      hatches: [],
      fillImages: [],
      strokeDashes: [{ name: "dash", attributes: { "draw:style": "round" } }],
    });
    expect(roundDash.shapeProperties?.outline?.customDash).toEqual([]);
    expect(roundDash.shapeProperties?.styleName).toBe("chart-style");
  });

  it("converts percentage and absolute dash lengths", async () => {
    const dash = (length: string, distance: string) => ({
      styles: [
        style({
          "draw:fill": "none",
          "draw:stroke": "dash",
          "draw:stroke-dash": "dash",
          "svg:stroke-width": "0.1cm",
          "svg:stroke-color": "#123456",
        }),
      ],
      gradients: [],
      hatches: [],
      fillImages: [],
      strokeDashes: [
        {
          name: "dash",
          attributes: {
            "draw:style": "rect",
            "draw:dots1": "1",
            "draw:dots1-length": length,
            "draw:distance": distance,
          },
        },
      ],
    });
    expect((await parse(dash("50%", "20%"))).shapeProperties?.outline?.customDash).toEqual([
      { d: "50%", sp: "20%" },
    ]);
    expect((await parse(dash("0.5cm", "0.2cm"))).shapeProperties?.outline?.customDash).toEqual([
      { d: "500%", sp: "200%" },
    ]);
  });

  it("omits broken fill references but reports missing media and dash width", async () => {
    const unresolved = await parse({
      styles: [style({ "draw:fill": "gradient", "draw:fill-gradient-name": "missing" })],
      gradients: [],
      hatches: [],
      fillImages: [],
      strokeDashes: [],
    });
    expect(unresolved.shapeProperties?.fill).toBeUndefined();
    await expect(
      parse({
        styles: [style({ "draw:fill": "bitmap", "draw:fill-image-name": "image" })],
        gradients: [],
        hatches: [],
        fillImages: [{ name: "image", attributes: { "xlink:href": "media/missing.png" } }],
        strokeDashes: [],
      }),
    ).rejects.toMatchObject({ reason: "missing-fill-image-part" });
    await expect(
      parse({
        styles: [style({ "draw:fill": "bitmap", "draw:fill-image-name": "image" })],
        gradients: [],
        hatches: [],
        fillImages: [{ name: "image", attributes: { "xlink:href": "media/missing.png" } }],
        strokeDashes: [],
      }),
    ).rejects.toMatchObject({ reason: "missing-fill-image-part" });
    await expect(
      parse({
        styles: [style({ "draw:stroke": "dash", "draw:stroke-dash": "dash" })],
        gradients: [],
        hatches: [],
        fillImages: [],
        strokeDashes: [
          { name: "dash", attributes: { "draw:dots1": "1", "draw:dots1-length": "1cm" } },
        ],
      }),
    ).rejects.toMatchObject({ reason: "missing-dash-basis" });
  });

  it("round-trips typed resources, references, and projection semantics", async () => {
    const resources: ChartStyleResources = {
      styles: [style({ "draw:fill": "gradient", "draw:fill-gradient-name": "gradient" })],
      gradients: [
        {
          name: "gradient",
          displayName: "Gradient",
          attributes: {
            "draw:style": "linear",
            "draw:angle": "30",
            "draw:start-color": "#ffffff",
            "draw:end-color": "#000000",
          },
        },
      ],
      hatches: [{ name: "hatch", attributes: { "draw:style": "single", "draw:color": "#000000" } }],
      fillImages: [
        { name: "image", attributes: { "xlink:href": "https://example.invalid/a.png" } },
      ],
      strokeDashes: [
        {
          name: "dash",
          attributes: { "draw:style": "rect", "draw:dots1": "1", "draw:dots1-length": "1cm" },
        },
      ],
    };
    const first = await parse(resources);
    const second = (await parseChartDocument(generateChartDocument({ chart: first }))).chart;
    expect(second.styleResources?.gradients).toEqual(resources.gradients);
    expect(second.shapeProperties?.styleName).toBe("chart-style");
    expect(second.shapeProperties?.fill).toEqual(first.shapeProperties?.fill);
  });
});

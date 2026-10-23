import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateChartDocument, parseChartDocument } from "./index";

const decoder = new TextDecoder();

describe("ODF chart inline decorations", () => {
  it("round-trips custom shapes and polylines", async () => {
    const options = {
      chart: {
        type: "column" as const,
        series: [{ name: "Series", values: [1, 2] }],
        odfShapes: [
          {
            kind: "customShape" as const,
            name: "Arrow",
            x: 100000,
            y: 200000,
            width: 300000,
            height: 400000,
            endCellAddress: "Sheet1.F12",
            endX: 50000,
            endY: 60000,
            text: "Chart & <shape>",
            enhancedGeometry: {
              type: "mso-spt32",
              enhancedPath: "M 0 0 L 21600 21600 N",
              viewBox: [0, 0, 21600, 21600] as const,
              mirrorHorizontal: false,
              mirrorVertical: false,
            },
          },
          {
            kind: "polyline" as const,
            name: "Top",
            points: [
              { x: 0, y: 100000 },
              { x: 200000, y: 0 },
            ],
            enhancedGeometry: { viewBox: [0, 0, 200000, 100000] as const },
          },
        ],
      },
    };
    const parsed = await parseChartDocument(generateChartDocument(options));
    const generated = decoder.decode(unzipSync(generateChartDocument(options))["content.xml"]!);
    expect(generated).toContain("Chart &amp; &lt;shape&gt;");
    expect(parsed.chart.odfShapes).toMatchObject(options.chart.odfShapes);
  });
});

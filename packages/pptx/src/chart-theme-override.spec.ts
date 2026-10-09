import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentationSync } from "./generate";
import { parsePresentationSync } from "./parse";
import type { PresentationOptions } from "./shared/file";

const RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";

describe("chart-owned theme overrides", () => {
  it("round-trips the typed override part with its source relationship", () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              chart: {
                type: "column",
                categories: ["A", "B"],
                series: [{ name: "Series", values: [1, 2] }],
                sourcePath: "ppt/charts/chart9.xml",
                sourceRelationships: [
                  {
                    relationshipType: RELATIONSHIP_TYPE,
                    target: "../theme/themeOverride9.xml",
                    rId: "rId1",
                  },
                ],
                themeOverride: { path: "ppt/theme/themeOverride9.xml", options: {} },
              },
            },
          ],
        },
      ],
    };

    const buffer = generatePresentationSync(options);
    const parsed = parsePresentationSync(buffer);
    const chart = parsed.slides?.[0]?.children?.[0];
    const files = unzipSync(buffer);

    expect(chart && "chart" in chart ? chart.chart.themeOverride : undefined).toEqual({
      options: {},
      path: "ppt/theme/themeOverride9.xml",
    });
    expect(files["ppt/theme/themeOverride9.xml"]).toBeDefined();
    expect(new TextDecoder().decode(files["ppt/charts/_rels/chart9.xml.rels"]!)).toContain(
      'Id="rId1"',
    );
  });
});

import { unzipSync, zipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generateDocumentSync } from "../../../generate";
import { parseDocumentSync } from "../../../parse";
import type { ChartExOptions } from "./chart-ex-run";

const chartExOptions = {
  chartData: {
    data: [
      {
        id: 0,
        dimensions: [
          {
            type: "val",
            formula: { text: "Sheet1!$A$2:$A$4" },
            levels: [
              {
                pointCount: 3,
                formatCode: "General",
                points: [
                  { index: 0, value: 12 },
                  { index: 1, value: -4 },
                  { index: 2, value: 30 },
                ],
              },
            ],
          },
        ],
      },
    ],
  },
  chart: {
    plotArea: {
      region: {
        series: [
          {
            layout: "waterfall",
            dataId: 0,
            layoutProperties: { subtotals: [1] },
          },
        ],
      },
      axes: [
        { id: 0, scaling: { kind: "category", gapWidth: 0.5 }, tickLabels: {} },
        { id: 1, scaling: { kind: "value" }, tickLabels: {} },
      ],
    },
  },
  transformation: { width: 5486400, height: 3200400 },
} satisfies ChartExOptions;

describe("DOCX ChartEx", () => {
  it("round-trips the typed chartEx part through canonical options", () => {
    const packed = generateDocumentSync({
      sections: [{ children: [{ paragraph: { children: [{ chartEx: chartExOptions }] } }] }],
    });
    const generated = unzipSync(packed);
    const part = new TextDecoder().decode(generated["word/charts/chartEx1.xml"] as Uint8Array);
    const rels = new TextDecoder().decode(generated["word/_rels/document.xml.rels"] as Uint8Array);
    const contentTypes = new TextDecoder().decode(generated["[Content_Types].xml"] as Uint8Array);
    expect(part).toContain('layoutId="waterfall"');
    expect(rels).toContain('Type="http://schemas.microsoft.com/office/2014/relationships/chartEx"');
    expect(contentTypes).toContain(
      'PartName="/word/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"',
    );

    const parsed = parseDocumentSync(packed);
    const child = parsed.sections?.[0]?.children?.[0];
    const paragraph = child as { paragraph?: { children?: unknown[] } };
    expect(paragraph.paragraph?.children?.[0]).toMatchObject({ chartEx: chartExOptions });

    const regenerated = unzipSync(generateDocumentSync(parsed));
    expect(regenerated["word/charts/chartEx1.xml"]).toEqual(generated["word/charts/chartEx1.xml"]);
  });

  it("parses ChartEx references by namespace-local name rather than prefix", () => {
    const generated = unzipSync(
      generateDocumentSync({
        sections: [{ children: [{ paragraph: { children: [{ chartEx: chartExOptions }] } }] }],
      }),
    );
    const documentXml = new TextDecoder().decode(generated["word/document.xml"]);
    generated["word/document.xml"] = new TextEncoder().encode(
      documentXml.replaceAll("<cx:", "<chart14:").replaceAll("</cx:", "</chart14:"),
    );
    const aliased = zipSync(generated);
    const parsed = parseDocumentSync(aliased);
    const child = parsed.sections?.[0]?.children?.[0];
    const paragraph = child as { paragraph?: { children?: unknown[] } };
    expect(paragraph.paragraph?.children?.[0]).toMatchObject({ chartEx: chartExOptions });
  });
});

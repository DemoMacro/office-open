// Demo: group-nested chart frames (wpg:graphicFrame), linked text boxes
// (wps:linkedTxbx + normalEastAsianFlow), and content part references
import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        { paragraph: { children: [{ bold: true, text: "Grouped Drawing Frames", size: 16 }] } },

        // 1. Group containing a chart frame and a text-box shape — Word emits
        //    charts inside groups as wpg:graphicFrame with cNvPr/cNvFrPr/xfrm.
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  transformation: { height: "8cm", width: "16cm" },
                  children: [
                    {
                      type: "chart",
                      chartOptions: {
                        type: "column",
                        categories: ["Q1", "Q2", "Q3", "Q4"],
                        series: [{ name: "2025", values: [140, 170, 210, 250] }],
                        title: "Quarterly Revenue",
                      },
                      transformation: {
                        pixels: { x: 0, y: 0 },
                        emus: { x: 5_486_400, y: 2_743_200 },
                      },
                    },
                    {
                      type: "wps",
                      data: {
                        children: ["Grouped caption shape"],
                      },
                      transformation: {
                        offset: { emus: { x: 0, y: 2_743_200 }, pixels: { x: 0, y: 0 } },
                        pixels: { x: 0, y: 0 },
                        emus: { x: 5_486_400, y: 914_400 },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Linked text box chain reference — the text lives in the linked
        //    part, so the shape carries id/seq instead of inline content.
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [],
                  linkedTextBox: { id: 1, sequence: 1 },
                  normalEastAsianFlow: true,
                  transformation: { height: "2cm", width: "6cm" },
                },
              },
            ],
          },
        },

        // 3. Content part reference (wp:contentPart) — opaque part via r:id
        {
          paragraph: {
            children: [
              {
                contentPart: {
                  referenceId: "rId1",
                  transformation: {
                    pixels: { x: 0, y: 0 },
                    emus: { x: 1_828_800, y: 914_400 },
                  },
                },
              },
            ],
          },
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/127-grouped-drawing-frames.docx", buffer);

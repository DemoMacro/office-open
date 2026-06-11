// Demo: Effect DAG container (CT_EffectContainer, a:effectDag) - DrawingML
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Effect DAG Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // 1. Simple effectDag with glow and shadow
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Glow + Outer Shadow",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [
                        {
                          bold: true,
                          color: "FFFFFF",
                          text: "Glow + Shadow",
                        },
                      ],
                    },
                  ],
                  effectDag: {
                    glow: { radius: 50800, color: { value: "00B0F0" } },
                    outerShadow: {
                      blurRadius: 76200,
                      color: { value: "000000" },
                      direction: 5400000,
                      distance: 38100,
                    },
                  },
                  fill: "0070C0",
                  transformation: { height: 80, width: 300 },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Alpha effects
        {
          paragraph: {
            children: [{ bold: true, text: "2. Luminance + Tint", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: ["Color Adjusted"],
                    },
                  ],
                  effectDag: {
                    luminance: { bright: 20000, contrast: -10000 },
                    tint: { hue: 2700000, amount: 30000 },
                  },
                  fill: "4472C4",
                  transformation: { height: 80, width: 300 },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Multiple effects with type attribute
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. Multiple Effects (sib type)",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [
                        {
                          bold: true,
                          color: "FFFFFF",
                          text: "Complex Effects",
                        },
                      ],
                    },
                  ],
                  effectDag: {
                    type: "sibling",
                    blur: { radius: 40000 },
                    softEdge: 25400,
                    outerShadow: {
                      color: { value: "000000" },
                      blurRadius: 50800,
                      direction: 2700000,
                      distance: 25400,
                    },
                  },
                  fill: {
                    type: "gradient",
                    options: {
                      shade: { angle: 5400000 },
                      stops: [
                        { color: { value: "1F4E79" }, position: 0 },
                        { color: { value: "2E75B6" }, position: 100000 },
                      ],
                    },
                  },
                  transformation: { height: 80, width: 300 },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 4. Effect DAG with container nesting
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "4. Nested Container (tree type)",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [
                        {
                          bold: true,
                          text: "Nested Effects",
                        },
                      ],
                    },
                  ],
                  effectDag: {
                    type: "tree",
                    containers: [
                      {
                        glow: { color: { value: "FF0000" } },
                        outerShadow: { color: { value: "000000" } },
                      },
                    ],
                  },
                  outline: {
                    color: { value: "C00000" },
                    type: "solidFill",
                    width: 25400,
                  },
                  fill: "F4B183",
                  transformation: { height: 80, width: 300 },
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);

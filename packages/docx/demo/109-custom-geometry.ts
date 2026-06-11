// Demo: Custom geometry (CT_CustomGeometry2D) - DrawingML
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Custom Geometry Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // 1. Simple triangle
        {
          paragraph: {
            children: [{ bold: true, text: "1. Triangle", size: 14 }],
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
                      children: ["Triangle"],
                    },
                  ],
                  customGeometry: {
                    pathList: [
                      {
                        w: 100000,
                        h: 100000,
                        commands: [
                          { command: "moveTo", point: { x: "50000", y: "0" } },
                          {
                            command: "lineTo",
                            point: { x: "100000", y: "100000" },
                          },
                          { command: "lineTo", point: { x: "0", y: "100000" } },
                          { command: "close" },
                        ],
                      },
                    ],
                  },
                  fill: "4472C4",
                  transformation: { height: 150, width: 200 },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Diamond
        {
          paragraph: {
            children: [{ bold: true, text: "2. Diamond", size: 14 }],
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
                      children: ["Diamond"],
                    },
                  ],
                  customGeometry: {
                    pathList: [
                      {
                        w: 100000,
                        h: 100000,
                        commands: [
                          { command: "moveTo", point: { x: "50000", y: "0" } },
                          {
                            command: "lineTo",
                            point: { x: "100000", y: "50000" },
                          },
                          {
                            command: "lineTo",
                            point: { x: "50000", y: "100000" },
                          },
                          { command: "lineTo", point: { x: "0", y: "50000" } },
                          { command: "close" },
                        ],
                      },
                    ],
                  },
                  outline: {
                    color: { value: "C00000" },
                    type: "solidFill",
                    width: 12700,
                  },
                  fill: "F4B183",
                  transformation: { height: 150, width: 200 },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Shape with arc
        {
          paragraph: {
            children: [{ bold: true, text: "3. Shape with Arc", size: 14 }],
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
                      children: ["Arc Shape"],
                    },
                  ],
                  customGeometry: {
                    pathList: [
                      {
                        w: 100000,
                        h: 100000,
                        commands: [
                          { command: "moveTo", point: { x: "0", y: "50000" } },
                          {
                            command: "arcTo",
                            heightRadius: "50000",
                            startAngle: "0",
                            sweepAngle: "5400000",
                            widthRadius: "50000",
                          },
                          { command: "lineTo", point: { x: "0", y: "50000" } },
                          { command: "close" },
                        ],
                      },
                    ],
                  },
                  fill: {
                    type: "gradient",
                    options: {
                      shade: { angle: 5400000 },
                      stops: [
                        { color: { value: "002060" }, position: 0 },
                        { color: { value: "00B0F0" }, position: 100000 },
                      ],
                    },
                  },
                  outline: {
                    color: { value: "002060" },
                    type: "solidFill",
                    width: 12700,
                  },
                  transformation: { height: 150, width: 200 },
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

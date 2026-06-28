import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Group Shape Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Group Shape Demo" },
            fill: "4472C4",
          },
        },
        {
          group: {
            x: "1.3cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "5.3cm",
            children: [
              {
                shape: {
                  x: "0.0cm",
                  y: "0.0cm",
                  width: "3.7cm",
                  height: "2.4cm",
                  textBody: { text: "Shape A" },
                  fill: "ED7D31",
                },
              },
              {
                shape: {
                  x: "4.2cm",
                  y: "0.0cm",
                  width: "3.7cm",
                  height: "2.4cm",
                  textBody: { text: "Shape B" },
                  fill: "70AD47",
                },
              },
              {
                shape: {
                  x: "0.0cm",
                  y: "2.9cm",
                  width: "7.9cm",
                  height: "2.4cm",
                  textBody: { text: "Shape C (wide)" },
                  fill: "5B9BD5",
                },
              },
            ],
          },
        },
        {
          group: {
            x: "10.6cm",
            y: "3.2cm",
            width: "6.6cm",
            height: "5.3cm",
            rotation: 10,
            children: [
              {
                shape: {
                  x: "0.0cm",
                  y: "0.0cm",
                  width: "6.6cm",
                  height: "5.3cm",
                  fill: "FFC000",
                },
              },
              {
                shape: {
                  x: "0.7cm",
                  y: "0.7cm",
                  width: "5.3cm",
                  height: "4.0cm",
                  textBody: { text: "Rotated Group" },
                  fill: "FFFFFF",
                },
              },
            ],
          },
        },
      ],
    },
    {
      background: { fill: "F2F2F2" },
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Nested Groups" },
            fill: "7030A0",
          },
        },
        {
          group: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "9.3cm",
            children: [
              {
                group: {
                  x: "0.0cm",
                  y: "0.0cm",
                  width: "7.4cm",
                  height: "4.2cm",
                  children: [
                    {
                      shape: {
                        x: "0.0cm",
                        y: "0.0cm",
                        width: "3.4cm",
                        height: "4.2cm",
                        textBody: { text: "Inner A" },
                        fill: "4472C4",
                      },
                    },
                    {
                      shape: {
                        x: "4.0cm",
                        y: "0.0cm",
                        width: "3.4cm",
                        height: "4.2cm",
                        textBody: { text: "Inner B" },
                        fill: "ED7D31",
                      },
                    },
                  ],
                },
              },
              {
                group: {
                  x: "7.9cm",
                  y: "0.0cm",
                  width: "7.9cm",
                  height: "9.3cm",
                  children: [
                    {
                      shape: {
                        x: "0.0cm",
                        y: "0.0cm",
                        width: "7.9cm",
                        height: "4.2cm",
                        textBody: { text: "Right Top" },
                        fill: "70AD47",
                      },
                    },
                    {
                      shape: {
                        x: "0.0cm",
                        y: "4.8cm",
                        width: "3.7cm",
                        height: "4.5cm",
                        textBody: { text: "RT Bot L" },
                        fill: "FFC000",
                      },
                    },
                    {
                      shape: {
                        x: "4.2cm",
                        y: "4.8cm",
                        width: "3.7cm",
                        height: "4.5cm",
                        textBody: { text: "RT Bot R" },
                        fill: "5B9BD5",
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);

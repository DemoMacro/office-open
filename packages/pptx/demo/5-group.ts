import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Group Shape Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Group Shape Demo" },
            fill: "4472C4",
          },
        },
        {
          group: {
            x: 50,
            y: 120,
            width: 300,
            height: 200,
            children: [
              {
                shape: {
                  x: 0,
                  y: 0,
                  width: 140,
                  height: 90,
                  textBody: { text: "Shape A" },
                  fill: "ED7D31",
                },
              },
              {
                shape: {
                  x: 160,
                  y: 0,
                  width: 140,
                  height: 90,
                  textBody: { text: "Shape B" },
                  fill: "70AD47",
                },
              },
              {
                shape: {
                  x: 0,
                  y: 110,
                  width: 300,
                  height: 90,
                  textBody: { text: "Shape C (wide)" },
                  fill: "5B9BD5",
                },
              },
            ],
          },
        },
        {
          group: {
            x: 400,
            y: 120,
            width: 250,
            height: 200,
            rotation: 10,
            children: [
              {
                shape: {
                  x: 0,
                  y: 0,
                  width: 250,
                  height: 200,
                  fill: "FFC000",
                },
              },
              {
                shape: {
                  x: 25,
                  y: 25,
                  width: 200,
                  height: 150,
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
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Nested Groups" },
            fill: "7030A0",
          },
        },
        {
          group: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            children: [
              {
                group: {
                  x: 0,
                  y: 0,
                  width: 280,
                  height: 160,
                  children: [
                    {
                      shape: {
                        x: 0,
                        y: 0,
                        width: 130,
                        height: 160,
                        textBody: { text: "Inner A" },
                        fill: "4472C4",
                      },
                    },
                    {
                      shape: {
                        x: 150,
                        y: 0,
                        width: 130,
                        height: 160,
                        textBody: { text: "Inner B" },
                        fill: "ED7D31",
                      },
                    },
                  ],
                },
              },
              {
                group: {
                  x: 300,
                  y: 0,
                  width: 300,
                  height: 350,
                  children: [
                    {
                      shape: {
                        x: 0,
                        y: 0,
                        width: 300,
                        height: 160,
                        textBody: { text: "Right Top" },
                        fill: "70AD47",
                      },
                    },
                    {
                      shape: {
                        x: 0,
                        y: 180,
                        width: 140,
                        height: 170,
                        textBody: { text: "RT Bot L" },
                        fill: "FFC000",
                      },
                    },
                    {
                      shape: {
                        x: 160,
                        y: 180,
                        width: 140,
                        height: 170,
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

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);

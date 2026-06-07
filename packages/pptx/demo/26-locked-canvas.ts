import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Locked Canvas Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 50,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Locked Canvas Demo", fontSize: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          lockedCanvas: {
            x: 100,
            y: 120,
            width: 600,
            height: 300,
            children: [
              {
                x: 20,
                y: 20,
                width: 250,
                height: 100,
                fill: "4472C4",
                textBody: {
                  children: [
                    {
                      properties: { alignment: "center", bullet: { type: "none" } },
                      children: [
                        {
                          text: "Locked Shape 1",
                          fontSize: 18,
                          bold: true,
                          fill: "FFFFFF",
                        },
                      ],
                    },
                  ],
                },
              },
              {
                x: 300,
                y: 20,
                width: 250,
                height: 100,
                fill: "ED7D31",
                textBody: {
                  children: [
                    {
                      properties: { alignment: "center", bullet: { type: "none" } },
                      children: [
                        {
                          text: "Locked Shape 2",
                          fontSize: 18,
                          bold: true,
                          fill: "FFFFFF",
                        },
                      ],
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

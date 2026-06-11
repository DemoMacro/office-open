import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Hyperlink Demo",
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
            textBody: { text: "Hyperlinks in PPTX" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: 50,
            y: 120,
            width: 600,
            height: 200,
            textBody: {
              children: [
                {
                  children: [
                    {
                      text: "Visit ",
                    },
                    {
                      text: "Google",
                      hyperlink: {
                        url: "https://www.google.com",
                        tooltip: "Go to Google",
                      },
                      bold: true,
                    },
                    {
                      text: " or ",
                    },
                    {
                      text: "GitHub",
                      hyperlink: { url: "https://github.com" },
                      italic: true,
                    },
                  ],
                },
                {
                  children: [
                    {
                      text: "Another link: ",
                    },
                    {
                      text: "Office Open XML Spec",
                      hyperlink: {
                        url: "https://www.iso.org/standard/71691.html",
                        tooltip: "ISO/IEC 29500",
                      },
                      fontSize: 14,
                    },
                  ],
                },
              ],
            },
            fill: "F2F2F2",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);

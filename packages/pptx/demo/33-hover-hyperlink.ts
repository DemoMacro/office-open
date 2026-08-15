import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Demonstrates a hover hyperlink (a:hlinkMouseOver) — a run hyperlink that
// fires on mouse-over instead of on click, distinct from a:hlinkClick.

const options: PresentationOptions = {
  title: "Hover hyperlink",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "16cm",
            height: "3cm",
            textBody: {
              paragraphs: [
                {
                  children: [
                    {
                      text: "Hover here for a tooltip",
                      mouseoverHyperlink: { url: "https://example.org", tooltip: "Hover tip" },
                    },
                  ],
                },
              ],
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/33-hover-hyperlink.pptx", buffer);

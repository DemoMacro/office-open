import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Demonstrates table cell diagonal borders (a:lnTlToBr / a:lnBlToTr) — used in
// financial and accounting tables to strike through a cell along one or both
// diagonals.

const options: PresentationOptions = {
  title: "Table diagonals",
  slides: [
    {
      children: [
        {
          table: {
            x: "2cm",
            y: "3cm",
            width: "16cm",
            height: "4cm",
            rows: [
              {
                cells: [
                  {
                    text: "Crossed",
                    borders: {
                      diagonalTopLeftToBottomRight: { color: "C00000", width: 12700 },
                      diagonalBottomLeftToTopRight: { color: "C00000", width: 12700 },
                    },
                  },
                  {
                    text: "One way",
                    borders: {
                      diagonalTopLeftToBottomRight: { color: "000000", width: 9525 },
                    },
                  },
                ],
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

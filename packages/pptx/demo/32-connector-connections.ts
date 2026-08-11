import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Demonstrates connector non-visual properties on slides: locking (a:cxnSpLocks)
// and endpoint connections (a:stCxn / a:endCxn) that glue a connector to shape
// connection sites. Connection ids reference shape cNvPr ids.

const options: PresentationOptions = {
  title: "Connector connections",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2cm",
            y: "3cm",
            width: "3cm",
            height: "3cm",
            geometry: "rect",
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "12cm",
            y: "3cm",
            width: "3cm",
            height: "3cm",
            geometry: "rect",
            fill: "ED7D31",
          },
        },
        {
          connector: {
            x1: "5cm",
            y1: "4.5cm",
            x2: "12cm",
            y2: "4.5cm",
            locking: { noAdjustHandles: true, noChangeShapeType: true },
            startConnection: { id: 1, index: 1 },
            endConnection: { id: 2, index: 3 },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);

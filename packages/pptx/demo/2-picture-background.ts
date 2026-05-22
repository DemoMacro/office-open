import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
  title: "Phase 2 Demo",
  creator: "Demo",
  slides: [
    {
      background: { fill: "F2F2F2" },
      children: [
        new Shape({
          x: 50,
          y: 50,
          width: 400,
          height: 100,
          text: "With Outline",
          geometry: "roundRect",
          fill: "FFFFFF",
          outline: { width: 25400, color: "4472C4", dashStyle: "dash" },
        }),
        new Shape({
          x: 500,
          y: 50,
          width: 400,
          height: 100,
          text: "Gradient Fill",
          fill: {
            type: "gradient",
            angle: 0,
            stops: [
              { position: 0, color: "4472C4" },
              { position: 100, color: "ED7D31" },
            ],
          },
        }),
      ],
    },
    {
      background: {
        fill: {
          type: "gradient",
          angle: 270,
          stops: [
            { position: 0, color: "1a1a2e" },
            { position: 100, color: "16213e" },
          ],
        },
      },
      children: [
        new Shape({
          x: 100,
          y: 200,
          width: 300,
          height: 200,
          text: "On Gradient BG",
          fill: "FFFFFF",
          outline: { width: 12700, color: "FFC000" },
        }),
      ],
    },
    {
      background: {
        fill: "1B2A4A",
        effects: {
          outerShadow: {
            blur: 50800,
            distance: 38100,
            direction: 2700000,
            color: "000000",
            alpha: 50,
          },
        },
      },
      children: [
        new Shape({
          x: 100,
          y: 100,
          width: 500,
          height: 80,
          text: "Background with Shadow Effect",
          fill: "4472C4",
        }),
      ],
    },
    {
      background: {
        fill: {
          type: "gradient",
          angle: 90,
          stops: [
            { position: 0, color: "2E4057" },
            { position: 100, color: "048A81" },
          ],
        },
        shadeToTitle: true,
      },
      children: [
        new Shape({
          x: 100,
          y: 100,
          width: 500,
          height: 80,
          text: "Background with shadeToTitle",
          fill: "FFFFFF",
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);

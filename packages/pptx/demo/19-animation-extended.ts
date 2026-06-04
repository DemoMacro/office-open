// Extended animation features: setBehavior, command, iterate.
//
// - setBehavior: instant property change (e.g., set visibility or opacity)
// - command: generic command (call/evt/verb)
// - iterate: text-level iteration (per-element, per-word, per-letter)

import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
  title: "Extended Animation Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        // Title
        new Shape({
          x: 50,
          y: 30,
          width: 500,
          height: 60,
          textBody: { text: "Extended Animation Demo" },
          fill: "4472C4",
          animation: { type: "fade", duration: 800 },
        }),

        // setBehavior: instantly set opacity to 0.5, then fade in
        new Shape({
          x: 50,
          y: 120,
          width: 300,
          height: 80,
          textBody: { text: "Set + Fade" },
          fill: "ED7D31",
          animation: {
            type: "fade",
            duration: 600,
            setBehavior: {
              attributeName: "style.opacity",
              toValue: "0.5",
              toType: "string",
            },
          },
        }),

        // Iterate: animate text per-character
        new Shape({
          x: 50,
          y: 230,
          width: 400,
          height: 80,
          textBody: { text: "Per-character animation" },
          fill: "70AD47",
          animation: {
            type: "fade",
            duration: 300,
            iterate: {
              type: "el",
              interval: 200,
              backwards: false,
            },
          },
        }),

        // Command: generic call command
        new Shape({
          x: 50,
          y: 340,
          width: 300,
          height: 80,
          textBody: { text: "Command" },
          fill: "FFC000",
          animation: {
            commandType: "call",
            command: "animate",
            duration: 1000,
            trigger: "withPrevious",
          },
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);

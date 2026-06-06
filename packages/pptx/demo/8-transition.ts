import * as fs from "fs";

import { Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
  title: "Transition Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Fade Transition" },
          fill: "4472C4",
        }),
      ],
      transition: { type: "fade", speed: "med" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Push Transition (Right)" },
          fill: "ED7D31",
        }),
      ],
      transition: { type: "push", direction: "right", speed: "slow" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Wipe Transition (Down)" },
          fill: "70AD47",
        }),
      ],
      transition: { type: "wipe", direction: "down" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Cover Transition (From Right)" },
          fill: "FFC000",
        }),
      ],
      transition: { type: "cover", direction: "right" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Split Transition" },
          fill: "5B9BD5",
        }),
      ],
      transition: { type: "split", orient: "horz", direction: "out" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Wheel Transition (4 spokes)" },
          fill: "BF8F00",
        }),
      ],
      transition: { type: "wheel", spokes: 4 },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Dissolve Transition" },
          fill: "7030A0",
        }),
      ],
      transition: { type: "dissolve", speed: "slow" },
    },
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Random Transition" },
          fill: "C00000",
        }),
      ],
      transition: { type: "random" },
    },
    // Transition with start sound
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Fade with Sound" },
          fill: "4472C4",
        }),
      ],
      transition: { type: "fade", speed: "med", startSound: { rId: "rId1", loop: true } },
    },
    // Transition with stop previous sound
    {
      children: [
        new Shape({
          x: 100,
          y: 150,
          width: 500,
          height: 100,
          textBody: { text: "Push (Stop Sound)" },
          fill: "ED7D31",
        }),
      ],
      transition: { type: "push", direction: "right", stopPreviousSound: true },
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);

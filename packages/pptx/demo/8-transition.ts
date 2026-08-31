import * as fs from "fs";
import * as path from "path";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Synthesized click tone (~1 s) — see assets/README.md.
const clickSound = new Uint8Array(
  fs.readFileSync(path.join(import.meta.dirname, "assets/transition-click.mp3")),
);

const options: PresentationOptions = {
  title: "Transition Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Fade Transition" },
            properties: {
              fill: "4472C4",
            },
          },
        },
      ],
      transition: { type: "fade", speed: "medium" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Push Transition (Right)" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
      ],
      transition: { type: "push", direction: "right", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Wipe Transition (Down)" },
            properties: {
              fill: "70AD47",
            },
          },
        },
      ],
      transition: { type: "wipe", direction: "down" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Cover Transition (From Right)" },
            properties: {
              fill: "FFC000",
            },
          },
        },
      ],
      transition: { type: "cover", direction: "right" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Split Transition" },
            properties: {
              fill: "5B9BD5",
            },
          },
        },
      ],
      transition: { type: "split", orient: "horizontal", direction: "out" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Wheel Transition (4 spokes)" },
            properties: {
              fill: "BF8F00",
            },
          },
        },
      ],
      transition: { type: "wheel", spokes: 4 },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Dissolve Transition" },
            properties: {
              fill: "7030A0",
            },
          },
        },
      ],
      transition: { type: "dissolve", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Random Transition" },
            properties: {
              fill: "C00000",
            },
          },
        },
      ],
      transition: { type: "random" },
    },
    // Transition with start sound
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Fade with Sound" },
            properties: {
              fill: "4472C4",
            },
          },
        },
      ],
      transition: {
        type: "fade",
        speed: "medium",
        startSound: { data: clickSound, type: "mp3", name: "click", loop: true },
      },
    },
    // Transition with stop previous sound
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "4.0cm",
            width: "13.2cm",
            height: "2.6cm",
            textBody: { text: "Push (Stop Sound)" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
      ],
      transition: { type: "push", direction: "right", stopPreviousSound: true },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/8-transition.pptx", buffer);

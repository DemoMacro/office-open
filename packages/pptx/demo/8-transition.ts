import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Transition Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Fade Transition" },
            fill: "4472C4",
          },
        },
      ],
      transition: { type: "fade", speed: "med" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Push Transition (Right)" },
            fill: "ED7D31",
          },
        },
      ],
      transition: { type: "push", direction: "right", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Wipe Transition (Down)" },
            fill: "70AD47",
          },
        },
      ],
      transition: { type: "wipe", direction: "down" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Cover Transition (From Right)" },
            fill: "FFC000",
          },
        },
      ],
      transition: { type: "cover", direction: "right" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Split Transition" },
            fill: "5B9BD5",
          },
        },
      ],
      transition: { type: "split", orient: "horz", direction: "out" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Wheel Transition (4 spokes)" },
            fill: "BF8F00",
          },
        },
      ],
      transition: { type: "wheel", spokes: 4 },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Dissolve Transition" },
            fill: "7030A0",
          },
        },
      ],
      transition: { type: "dissolve", speed: "slow" },
    },
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Random Transition" },
            fill: "C00000",
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
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Fade with Sound" },
            fill: "4472C4",
          },
        },
      ],
      transition: { type: "fade", speed: "med", startSound: { rId: "rId1", loop: true } },
    },
    // Transition with stop previous sound
    {
      children: [
        {
          shape: {
            x: 100,
            y: 150,
            width: 500,
            height: 100,
            textBody: { text: "Push (Stop Sound)" },
            fill: "ED7D31",
          },
        },
      ],
      transition: { type: "push", direction: "right", stopPreviousSound: true },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);

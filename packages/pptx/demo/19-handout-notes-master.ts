// Parameterized Handout Master and Notes Master with custom options.
// Also demonstrates SmartArt with explicit color and style IDs.

import { writeFileSync } from "node:fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  includeHandoutMaster: true,
  includeNotesMaster: true,
  handoutMasterOptions: {
    colorMap: {
      bg1: "lt1",
      tx1: "dk1",
      accent1: "accent1",
    },
    headerFooter: {
      date: true,
      header: true,
      footer: true,
      slideNumber: true,
    },
  },
  notesMasterOptions: {
    colorMap: {
      bg1: "lt1",
      tx1: "dk1",
    },
    headerFooter: {
      date: true,
      slideNumber: true,
    },
    notesStyle: [
      { fontSize: 1400, marginLeft: 0, alignment: "l" },
      { fontSize: 1200, marginLeft: 457200 },
      { fontSize: 1200, marginLeft: 914400 },
    ],
  },
  slides: [
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 60,
            textBody: { text: "Parameterized Master Demo" },
            fill: "4472C4",
          },
        },
      ],
      notes: "This slide has notes that render using the custom notes master style.",
    },
    {
      children: [
        {
          smartart: {
            x: 50,
            y: 130,
            width: 600,
            height: 250,
            nodes: [{ text: "Step 1" }, { text: "Step 2" }, { text: "Step 3" }],
            layout: "default",
            style: "simple1",
            color: "colorful1",
          },
        },
      ],
      notes: "Second slide notes.",
    },
  ],
};

const buffer = await generate(options);
writeFileSync("My Presentation.pptx", buffer);

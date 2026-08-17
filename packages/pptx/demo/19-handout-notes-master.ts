// Parameterized Handout Master and Notes Master with custom options.
// Also demonstrates SmartArt with explicit color and style IDs.

import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  includeHandoutMaster: true,
  includeNotesMaster: true,
  handoutMasterOptions: {
    colorMapping: {
      background1: "light1",
      text1: "dark1",
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
    colorMapping: {
      background1: "light1",
      text1: "dark1",
    },
    headerFooter: {
      date: true,
      slideNumber: true,
    },
    notesStyle: {
      levels: [
        { alignment: "l", marginIndent: 0, defaultRun: { size: 14 } },
        { marginIndent: 457200, defaultRun: { size: 12 } },
        { marginIndent: 914400, defaultRun: { size: 12 } },
      ],
    },
  },
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "1.6cm",
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
            x: "1.3cm",
            y: "3.4cm",
            width: "15.9cm",
            height: "6.6cm",
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

const buffer = await generatePresentation(options);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/19-handout-notes-master.pptx", buffer);

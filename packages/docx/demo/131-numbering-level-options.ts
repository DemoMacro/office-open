// Numbering with level restart, legacy spacing/indent, and abstract numbering options.

import { writeFileSync } from "node:fs";

import { AlignmentType, LevelFormat, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  numbering: {
    config: [
      {
        reference: "decimal-with-restart",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            start: 1,
            // Restart numbering at level 1
            lvlRestart: 0,
            // Legacy compatibility spacing and indent
            legacy: { space: 720, indent: 360 },
            style: {
              paragraph: {
                indent: { left: 720, hanging: 360 },
              },
            },
          },
          {
            level: 1,
            format: LevelFormat.LOWER_LETTER,
            text: "%2)",
            alignment: AlignmentType.LEFT,
            start: 1,
            lvlRestart: 0,
            legacy: { space: 360, indent: 180 },
            style: {
              paragraph: {
                indent: { left: 1440, hanging: 360 },
              },
            },
          },
        ],
      },
    ],
  },

  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Numbering with Level Restart and Legacy Options",
                size: 16,
              },
            ],
          },
        },

        {
          paragraph: {
            numbering: { reference: "decimal-with-restart", level: 0 },
            children: ["First item"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "decimal-with-restart", level: 1 },
            children: ["Sub-item a"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "decimal-with-restart", level: 1 },
            children: ["Sub-item b"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "decimal-with-restart", level: 0 },
            children: ["Second item"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "decimal-with-restart", level: 1 },
            children: ["Sub-item a (restarted)"],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);

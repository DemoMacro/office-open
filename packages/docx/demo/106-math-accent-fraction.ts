// Demo: New OOXML features - Math accent, Fraction types

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        // ===== Math Accent =====
        {
          paragraph: {
            children: [{ text: "Math Accent Demo", bold: true, size: 14 }],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  equations: [
                    [
                      { accent: { children: ["x"] } },
                      " + ",
                      {
                        accent: {
                          accentCharacter: "̃",
                          children: ["y"],
                        },
                      },
                      " + ",
                      {
                        accent: {
                          accentCharacter: "̇",
                          children: ["z"],
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  equations: [
                    [
                      {
                        accent: {
                          accentCharacter: "⃗",
                          children: ["AB"],
                        },
                      },
                      " = ",
                      {
                        accent: {
                          accentCharacter: "̅",
                          children: ["CD"],
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },

        // ===== Fraction Types =====
        {
          paragraph: {
            children: [{ text: "Fraction Types Demo", bold: true, size: 14 }],
          },
        },
        {
          paragraph: {
            children: [
              "Standard bar fraction: ",
              {
                math: {
                  equations: [
                    [
                      {
                        fraction: {
                          numerator: ["a"],
                          denominator: ["b"],
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Skewed fraction: ",
              {
                math: {
                  equations: [
                    [
                      {
                        fraction: {
                          numerator: ["a"],
                          denominator: ["b"],
                          fractionType: "skw",
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Linear fraction: ",
              {
                math: {
                  equations: [
                    [
                      {
                        fraction: {
                          numerator: ["a"],
                          denominator: ["b"],
                          fractionType: "lin",
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "No-bar fraction: ",
              {
                math: {
                  equations: [
                    [
                      {
                        fraction: {
                          numerator: ["a"],
                          denominator: ["b"],
                          fractionType: "noBar",
                        },
                      },
                    ],
                  ],
                },
              },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/106-math-accent-fraction.docx", buffer);
